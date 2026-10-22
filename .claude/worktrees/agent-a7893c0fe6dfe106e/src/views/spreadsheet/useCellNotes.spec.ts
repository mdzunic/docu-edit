// SPEC-135's useCellNotes: the indicators and the active cell's description,
// the hover popover after NOTE_HOVER_DELAY_MS (and its grace, and
// sheet-thread-open), the shown boxes and View ▸ Show all notes, the note
// editor (Shift+F2 from the grid, a new note or the cell's own, Escape /
// blur committing, unchanged changing nothing, emptied deleting), Delete
// note and Show/Hide note — each ONE labelled undo step that undo restores —
// and notes off on a CSV workbook; plus the pure views and the placement
// behind them (cellNoteViews, sheetNoteLayout).
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { computed, effectScope, ref, shallowRef, triggerRef, type EffectScope } from "vue";
import { NOTE_DEFAULT_BOX, NOTE_HOVER_DELAY_MS, NOTE_UNDO_LABELS, NOTES_KEPT_AS_READ_TITLE } from "../../lib/spreadsheet/annotations";
import type { CellNote, CellThread, Sheet, SpreadsheetFormat, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { buildMergeIndex } from "../../lib/spreadsheet/merges";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { annotationDescription, annotationMarks, noteBoxViews, notePopoverView, repliesLabel } from "./cellNoteViews";
import { NOTE_POPOVER_WIDTH_PX, noteBoxPlacement, noteGridHooks, notePopoverPlacement } from "./sheetNoteLayout";
import { NOTE_HIDE_GRACE_MS, useCellNotes } from "./useCellNotes";
import { useSpreadsheetUndo } from "./useSpreadsheetUndo";

const scopes: EffectScope[] = [];
afterEach(() => {
  scopes.splice(0).forEach((scope) => scope.stop());
  vi.useRealTimers();
  document.body.innerHTML = "";
});
beforeEach(() => localStorage.clear());

const B2_NOTE: CellNote = {
  noteId: "n-B2",
  ref: { row: 1, col: 1 },
  author: "Alex Kim",
  text: "Alex Kim:\nQ3 figure, unaudited.",
  runs: [{ text: "Alex Kim:", font: { bold: true } }, { text: "\nQ3 figure, unaudited." }],
  source: { xlsxComment: "<comment/>" },
};
const B4_NOTE: CellNote = { noteId: "n-B4", ref: { row: 3, col: 1 }, author: "Alex Kim", text: "Always shown", visible: true, box: { widthPx: 200, heightPx: 100 } };
const D2_THREAD: CellThread = {
  threadId: "{T1}",
  ref: { row: 1, col: 3 },
  resolved: false,
  messages: [
    { messageId: "{T1}", personId: "p1", author: "Sam Lee", text: "Where does this number come from?", createdAt: "2026-03-01T09:00:00Z" },
    { messageId: "{M2}", personId: "p2", author: "Alex Kim", text: "From the Q3 ledger.", createdAt: "2026-03-01T10:00:00Z" },
    { messageId: "{M3}", personId: "p1", author: "Sam Lee", text: "Thanks, that helps.", createdAt: "2026-03-01T11:00:00Z" },
  ],
};
const E2_THREAD: CellThread = {
  threadId: "{T2}",
  ref: { row: 1, col: 4 },
  resolved: true,
  messages: [{ messageId: "{T2}", personId: "p2", author: "Alex Kim", text: "Rounded to the nearest hundred.", createdAt: "2026-03-02T09:00:00Z" }],
};

function workbook(format: SpreadsheetFormat = "xlsx"): SpreadsheetModel {
  const model = createBlankWorkbook();
  model.format = format;
  const base = model.sheets[0]!;
  const summary: Sheet = { ...base, id: "s1", name: "Summary", notes: [structuredClone(B2_NOTE), structuredClone(B4_NOTE)], threads: [structuredClone(D2_THREAD), structuredClone(E2_THREAD)] };
  const data: Sheet = { ...base, id: "s2", name: "Data", rows: [] };
  model.sheets = [summary, data];
  return model;
}

function setup(format: SpreadsheetFormat = "xlsx", blocked?: () => string) {
  const wb = shallowRef<SpreadsheetModel | null>(workbook(format));
  const activeIndex = ref(0);
  const activeSheet = computed(() => wb.value?.sheets[activeIndex.value] ?? null);
  const active = ref({ row: 0, col: 0 });
  const selection = computed(() => ({ active: active.value, ranges: [{ top: active.value.row, left: active.value.col, bottom: active.value.row, right: active.value.col }], activeRangeIndex: 0 }));
  const busy = ref(false);
  const openThread = vi.fn();
  const focusGrid = vi.fn();
  const scope = effectScope();
  scopes.push(scope);
  const { undo, notes } = scope.run(() => {
    const stack = useSpreadsheetUndo({ model: wb, recalc: (_keys, done) => done([]), onApplied: () => triggerRef(wb), activateSheet: () => {} });
    return {
      undo: stack,
      notes: useCellNotes({ book: { workbook: wb, activeSheet }, undo: stack, selection: { selection }, busy: () => busy.value, openThread, focusGrid, author: () => "Pat Doe", ...(blocked ? { blocked } : {}) }),
    };
  })!;
  const sheet = (): Sheet => wb.value!.sheets[0]!;
  return { wb, activeIndex, active, busy, openThread, focusGrid, undo, notes, sheet };
}

/** A keydown dispatched from inside a grid element. */
function gridKey(init: KeyboardEventInit): KeyboardEvent {
  const grid = document.createElement("div");
  grid.dataset.testid = "spreadsheet-grid";
  const editor = document.createElement("div");
  grid.append(editor);
  document.body.append(grid);
  const event = new KeyboardEvent("keydown", { bubbles: true, cancelable: true, ...init });
  editor.dispatchEvent(event);
  return event;
}

describe("useCellNotes — indicators and the active cell's description (SPEC-135)", () => {
  it("marks a note's cell, an open thread's and a resolved thread's; nothing elsewhere", () => {
    const { notes } = setup();
    expect(notes.marksAt(1, 1)).toEqual({ note: true, thread: null });
    expect(notes.marksAt(1, 3)).toEqual({ note: false, thread: "open" });
    expect(notes.marksAt(1, 4)).toEqual({ note: false, thread: "resolved" });
    expect(notes.marksAt(0, 0)).toBeNull();
  });

  it("describes the active cell's note or thread in plain text, and nothing for a bare cell", () => {
    const { notes, active } = setup();
    active.value = { row: 1, col: 1 };
    expect(notes.activeDescription.value).toBe("Note by Alex Kim: Alex Kim:\nQ3 figure, unaudited.");
    active.value = { row: 1, col: 3 };
    expect(notes.activeDescription.value).toBe("Comment by Sam Lee: Where does this number come from? 2 replies");
    active.value = { row: 1, col: 4 };
    expect(notes.activeDescription.value).toBe("Comment by Alex Kim (resolved): Rounded to the nearest hundred.");
    active.value = { row: 0, col: 0 };
    expect(notes.activeDescription.value).toBe("");
  });
});

describe("useCellNotes — hovering (SPEC-135)", () => {
  it("shows a note's popover only after NOTE_HOVER_DELAY_MS, and hides it a grace after the pointer leaves", () => {
    vi.useFakeTimers();
    const { notes } = setup();
    notes.hoverCell({ row: 1, col: 1 });
    vi.advanceTimersByTime(NOTE_HOVER_DELAY_MS - 1);
    expect(notes.popover.value).toBeNull();
    vi.advanceTimersByTime(1);
    expect(notes.popover.value).toMatchObject({ ref: { row: 1, col: 1 }, note: { author: "Alex Kim", text: "Alex Kim:\nQ3 figure, unaudited." }, thread: null });
    notes.hoverCell(null);
    vi.advanceTimersByTime(NOTE_HIDE_GRACE_MS - 1);
    expect(notes.popover.value).not.toBeNull();
    // The pointer reaching the popover keeps it.
    notes.hoverPopover(true);
    vi.advanceTimersByTime(NOTE_HIDE_GRACE_MS * 2);
    expect(notes.popover.value).not.toBeNull();
    notes.hoverPopover(false);
    vi.advanceTimersByTime(NOTE_HIDE_GRACE_MS);
    expect(notes.popover.value).toBeNull();
  });

  it("a thread's popover carries its opening message and reply count; sheet-thread-open opens the panel on it", () => {
    vi.useFakeTimers();
    const { notes, openThread } = setup();
    notes.hoverCell({ row: 1, col: 3 });
    vi.advanceTimersByTime(NOTE_HOVER_DELAY_MS);
    expect(notes.popover.value?.thread).toEqual({ threadId: "{T1}", author: "Sam Lee", text: "Where does this number come from?", replies: 2, resolved: false });
    notes.openPopoverThread();
    expect(openThread).toHaveBeenCalledWith("s1", "{T1}");
    expect(notes.popover.value).toBeNull();
  });

  it("a bare cell, a shown note's cell and a busy grid start no popover; leaving before the delay cancels it", () => {
    vi.useFakeTimers();
    const { notes, busy } = setup();
    notes.hoverCell({ row: 0, col: 0 });
    notes.hoverCell({ row: 3, col: 1 });
    vi.advanceTimersByTime(NOTE_HOVER_DELAY_MS);
    // B4's note is visible: it shows as a box, never as a popover.
    expect(notes.popover.value).toBeNull();
    notes.hoverCell({ row: 1, col: 1 });
    vi.advanceTimersByTime(NOTE_HOVER_DELAY_MS / 2);
    notes.hoverCell(null);
    vi.advanceTimersByTime(NOTE_HOVER_DELAY_MS);
    expect(notes.popover.value).toBeNull();
    busy.value = true;
    notes.hoverCell({ row: 1, col: 1 });
    vi.advanceTimersByTime(NOTE_HOVER_DELAY_MS);
    expect(notes.popover.value).toBeNull();
  });
});

describe("useCellNotes — boxes and Show all notes (SPEC-135)", () => {
  it("shows the visible note at its box size; Show all notes shows every note, session-only, and records nothing", () => {
    const { notes, undo, wb } = setup();
    expect(notes.boxes.value.map((box) => [box.noteId, box.widthPx, box.heightPx])).toEqual([["n-B4", 200, 100]]);
    const before = wb.value!.sheets;
    notes.toggleShowAll();
    expect(notes.showAll.value).toBe(true);
    expect(notes.boxes.value.map((box) => [box.noteId, box.widthPx, box.heightPx])).toEqual([
      ["n-B2", NOTE_DEFAULT_BOX.widthPx, NOTE_DEFAULT_BOX.heightPx],
      ["n-B4", 200, 100],
    ]);
    expect(wb.value!.sheets).toBe(before);
    expect(undo.canUndo.value).toBe(false);
    expect(localStorage.length).toBe(0);
  });
});

describe("useCellNotes — the note editor and the note commands (SPEC-135)", () => {
  it("Shift+F2 from the grid opens the editor on the active cell; Escape-style commit adds the note as one 'Add note' step by the identity", () => {
    const { notes, active, undo, sheet, focusGrid } = setup();
    active.value = { row: 3, col: 3 };
    const event = gridKey({ key: "F2", code: "F2", shiftKey: true });
    expect(event.defaultPrevented).toBe(true);
    expect(notes.editor.value).toMatchObject({ sheetId: "s1", ref: { row: 3, col: 3 }, noteId: null, ...NOTE_DEFAULT_BOX });
    expect(notes.draft.value).toBe("");
    notes.input("Check with finance");
    notes.commitEditor(true);
    expect(notes.editor.value).toBeNull();
    expect(focusGrid).toHaveBeenCalled();
    const added = sheet().notes!.find((note) => note.ref.row === 3 && note.ref.col === 3);
    expect(added).toMatchObject({ author: "Pat Doe", text: "Check with finance" });
    expect(sheet().annotationsDirty).toBe(true);
    expect(undo.undoLabel.value).toBe(NOTE_UNDO_LABELS.add);
    expect(notes.marksAt(3, 3)).toEqual({ note: true, thread: null });
    undo.undo();
    expect(sheet().notes!.some((note) => note.ref.row === 3 && note.ref.col === 3)).toBe(false);
  });

  it("Shift+F2 on a noted cell edits its note; an unchanged draft changes nothing (the rich runs stay); an edit is 'Edit note'", () => {
    const { notes, active, undo, sheet } = setup();
    active.value = { row: 1, col: 1 };
    gridKey({ key: "F2", code: "F2", shiftKey: true });
    expect(notes.editor.value?.noteId).toBe("n-B2");
    expect(notes.draft.value).toBe("Alex Kim:\nQ3 figure, unaudited.");
    notes.commitEditor(false);
    expect(sheet().notes![0]!.runs).toHaveLength(2);
    expect(undo.canUndo.value).toBe(false);
    notes.openEditor();
    notes.input("Audited now.");
    notes.commitEditor(false);
    expect(sheet().notes![0]).toMatchObject({ noteId: "n-B2", text: "Audited now." });
    expect(sheet().notes![0]!.runs).toBeUndefined();
    expect(sheet().notes![0]!.source).toBeUndefined();
    expect(undo.undoLabel.value).toBe(NOTE_UNDO_LABELS.edit);
  });

  it("an emptied note deletes it ('Delete note'); a new note left empty adds nothing", () => {
    const { notes, active, undo, sheet } = setup();
    active.value = { row: 1, col: 1 };
    notes.openEditor();
    notes.input("   ");
    notes.commitEditor(false);
    expect(sheet().notes!.map((note) => note.noteId)).toEqual(["n-B4"]);
    expect(undo.undoLabel.value).toBe(NOTE_UNDO_LABELS.delete);
    active.value = { row: 9, col: 9 };
    notes.openEditor();
    notes.commitEditor(false);
    expect(undo.undoLabel.value).toBe(NOTE_UNDO_LABELS.delete);
  });

  it("Delete note and Show note / Hide note act on the active cell, each one labelled step", () => {
    const { notes, active, undo, sheet } = setup();
    active.value = { row: 1, col: 1 };
    notes.toggleActiveVisible();
    expect(sheet().notes![0]!.visible).toBe(true);
    expect(undo.undoLabel.value).toBe(NOTE_UNDO_LABELS.show);
    notes.toggleActiveVisible();
    expect(sheet().notes![0]!.visible).toBeUndefined();
    expect(undo.undoLabel.value).toBe(NOTE_UNDO_LABELS.hide);
    notes.deleteActive();
    expect(sheet().notes!.map((note) => note.noteId)).toEqual(["n-B4"]);
    expect(undo.undoLabel.value).toBe(NOTE_UNDO_LABELS.delete);
    undo.undo();
    expect(sheet().notes!.map((note) => note.noteId)).toEqual(["n-B2", "n-B4"]);
  });

  it("an .ods note marks the sheet dirty; a CSV workbook offers no notes; nothing opens while a cell is edited", () => {
    const ods = setup("ods");
    ods.active.value = { row: 0, col: 0 };
    ods.notes.openEditor();
    ods.notes.input("Hello");
    ods.notes.commitEditor(false);
    expect(ods.sheet().dirty).toBe(true);
    expect(ods.sheet().annotationsDirty).toBeUndefined();
    const csv = setup("csv");
    expect(csv.notes.canAnnotate.value).toBe(false);
    expect(csv.notes.openEditor()).toBe(false);
    // Only the busy view's listener from here on.
    scopes.splice(0).forEach((scope) => scope.stop());
    const busy = setup();
    busy.busy.value = true;
    expect(busy.notes.openEditor()).toBe(false);
    expect(gridKey({ key: "F2", code: "F2", shiftKey: true }).defaultPrevented).toBe(false);
  });

  it("a sheet whose source's notes stay as read (AC-11) offers no note: disabled with NOTES_KEPT_AS_READ_TITLE, neither the editor nor Shift+F2 opens", () => {
    const reason = ref(NOTES_KEPT_AS_READ_TITLE);
    const kept = setup("xlsx", () => reason.value);
    expect(kept.notes.canAnnotate.value).toBe(false);
    expect(kept.notes.noteDisabledReason.value).toBe(NOTES_KEPT_AS_READ_TITLE);
    expect(kept.notes.openEditor()).toBe(false);
    expect(gridKey({ key: "F2", code: "F2", shiftKey: true }).defaultPrevented).toBe(false);
    expect(kept.notes.editor.value).toBeNull();
    reason.value = "";
    expect(kept.notes.canAnnotate.value).toBe(true);
    expect(kept.notes.noteDisabledReason.value).toBe("");
    // A CSV workbook has no reason to give: it never holds notes.
    expect(setup("csv", () => NOTES_KEPT_AS_READ_TITLE).notes.noteDisabledReason.value).toBe("");
  });

  it("switching sheets commits the open editor", () => {
    const { notes, active, activeIndex, sheet } = setup();
    active.value = { row: 5, col: 0 };
    notes.openEditor();
    notes.input("Before switching");
    activeIndex.value = 1;
    return Promise.resolve().then(() => {
      expect(notes.editor.value).toBeNull();
      expect(sheet().notes!.some((note) => note.text === "Before switching")).toBe(true);
    });
  });
});

describe("cellNoteViews and sheetNoteLayout (SPEC-135)", () => {
  it("builds the marks, the popover, the boxes and the description from a sheet", () => {
    const sheet = workbook().sheets[0]!;
    expect([...annotationMarks(sheet).keys()]).toEqual(["1:1", "3:1", "1:3", "1:4"]);
    expect(notePopoverView(sheet, { row: 3, col: 1 }, false)).toBeNull();
    expect(notePopoverView(sheet, { row: 1, col: 1 }, true)).toBeNull();
    expect(notePopoverView(sheet, { row: 1, col: 4 }, false)?.thread).toMatchObject({ replies: 0, resolved: true });
    expect(noteBoxViews(sheet, true, "n-B2").map((box) => box.noteId)).toEqual(["n-B4"]);
    expect(annotationDescription(undefined, undefined)).toBe("");
    expect([repliesLabel(0), repliesLabel(1), repliesLabel(2)]).toEqual(["0 replies", "1 reply", "2 replies"]);
  });

  it("places a box 15 px right of its cell and 10 px below the row above's top, and the popover beside the cell, flipped at the right edge", () => {
    const anchor = { cell: { left: 100, top: 40, width: 64, height: 20 }, above: 20 };
    expect(noteBoxPlacement(anchor, { widthPx: 144, heightPx: 79 })).toEqual({ left: 179, top: 30, width: 144, height: 79 });
    expect(noteBoxPlacement({ cell: { left: 0, top: 0, width: 64, height: 20 }, above: 0 }, NOTE_DEFAULT_BOX).top).toBe(10);
    expect(notePopoverPlacement(anchor.cell, { width: 1000 })).toEqual({ left: 168, top: 40 });
    expect(notePopoverPlacement({ left: 700, top: 40, width: 64, height: 20 }, { width: 800 })).toEqual({ left: 700 - 4 - NOTE_POPOVER_WIDTH_PX, top: 40 });
  });

  it("noteGridHooks: the anchor of a merged cell is its merge's; the pointer reports a cell by its data-ref, a popover, or nothing", () => {
    const hoverCell = vi.fn();
    const hoverPopover = vi.fn();
    const rectOf = vi.fn((range: { top: number; left: number; bottom: number; right: number }) => ({ left: range.left * 64, top: range.top * 20, width: (range.right - range.left + 1) * 64, height: (range.bottom - range.top + 1) * 20 }));
    const merges = buildMergeIndex([{ top: 1, left: 1, bottom: 2, right: 2 }]);
    const hooks = noteGridHooks({
      notes: () => ({ hoverCell, hoverPopover }),
      rectOf,
      metrics: () => ({ rowTop: (row: number) => row * 20 }),
      merges: () => merges,
    });
    expect(hooks.noteAnchor({ row: 0, col: 0 })).toEqual({ cell: { left: 0, top: 0, width: 64, height: 20 }, above: 0 });
    expect(hooks.noteAnchor({ row: 1, col: 1 })).toEqual({ cell: { left: 64, top: 20, width: 128, height: 40 }, above: 20 });
    const cell = document.createElement("div");
    cell.className = "sheet-cell";
    cell.dataset.ref = "D4";
    const inner = document.createElement("span");
    cell.append(inner);
    hooks.notePointer.over({ target: inner } as unknown as Event);
    expect(hoverCell).toHaveBeenLastCalledWith(expect.objectContaining({ row: 3, col: 3 }));
    const popover = document.createElement("div");
    popover.dataset.notePopover = "true";
    hooks.notePointer.over({ target: popover } as unknown as Event);
    expect(hoverPopover).toHaveBeenCalledWith(true);
    hooks.notePointer.leave();
    expect(hoverCell).toHaveBeenLastCalledWith(null);
  });
});
