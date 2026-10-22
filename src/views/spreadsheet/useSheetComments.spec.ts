// SPEC-135's useSheetComments: the panel's groups (by sheet in tab order,
// then (row, col); the unresolved filter; an .ods workbook's notes instead),
// the identity, New comment (the composer on the active cell, the cell's own
// thread when it has one, Ctrl+Alt+M from the grid, off in an .ods workbook
// with ODS_NOTES_ONLY_TITLE), reply / resolve / reopen / edit / delete —
// edit and delete on the identity's own messages only — each ONE labelled
// undo step that undo restores, the pane toggle, opening on a thread, and
// activating a thread selecting its sheet and cell.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { computed, effectScope, nextTick, ref, shallowRef, triggerRef, type EffectScope } from "vue";
import { COMMENT_UNDO_LABELS, ODS_NOTES_ONLY_TITLE } from "../../lib/spreadsheet/annotations";
import type { CellThread, Sheet, SpreadsheetFormat, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { createBlankWorkbook } from "../../lib/spreadsheet/sheets";
import { messageTime, useSheetComments } from "./useSheetComments";
import { useSheetSidePanel } from "./useSheetSidePanel";
import type { SheetViewFields } from "./useSheetViewState";
import { useSpreadsheetUndo } from "./useSpreadsheetUndo";

const scopes: EffectScope[] = [];
afterEach(() => {
  scopes.splice(0).forEach((scope) => scope.stop());
  document.body.innerHTML = "";
});
beforeEach(() => localStorage.clear());

const thread = (threadId: string, row: number, col: number, resolved: boolean, ...messages: [string, string][]): CellThread => ({
  threadId,
  ref: { row, col },
  resolved,
  messages: messages.map(([author, text], i) => ({ messageId: i === 0 ? threadId : `${threadId}-${i}`, personId: `p-${author}`, author, text, createdAt: "2026-03-01T09:00:00Z" })),
});

function workbook(format: SpreadsheetFormat): SpreadsheetModel {
  const model = createBlankWorkbook();
  model.format = format;
  const base = model.sheets[0]!;
  const summary: Sheet = {
    ...base,
    id: "s1",
    name: "Summary",
    threads: [
      thread("{D2}", 1, 3, false, ["Sam Lee", "Where does this number come from?"], ["Alex Kim", "From the Q3 ledger."], ["You", "Thanks."]),
      thread("{E2}", 1, 4, true, ["Alex Kim", "Rounded to the nearest hundred."]),
    ],
    notes: [{ noteId: "n-B2", ref: { row: 1, col: 1 }, author: "Alex Kim", text: "Q3 figure, unaudited.", createdAt: "2026-03-01T09:00:00" }],
  };
  const data: Sheet = { ...base, id: "s2", name: "Data", rows: [], threads: [thread("{A1}", 0, 0, false, ["You", "Header?"])] };
  model.sheets = [summary, data];
  model.persons = [{ personId: "p-Sam Lee", displayName: "Sam Lee" }];
  return model;
}

function setup(format: SpreadsheetFormat = "xlsx") {
  const wb = shallowRef<SpreadsheetModel | null>(workbook(format));
  const activeIndex = ref(0);
  const activeSheet = computed(() => wb.value?.sheets[activeIndex.value] ?? null);
  const fields = shallowRef<Record<string, SheetViewFields>>({});
  const viewState = { update: (sheetId: string, patch: SheetViewFields) => void (fields.value = { ...fields.value, [sheetId]: { ...fields.value[sheetId], ...patch } }) };
  const active = ref({ row: 0, col: 0 });
  const selection = computed(() => ({ active: active.value, ranges: [{ top: active.value.row, left: active.value.col, bottom: active.value.row, right: active.value.col }], activeRangeIndex: 0 }));
  const cell = { selectSheet: vi.fn((index: number) => void (activeIndex.value = index)), revealTick: ref(0) };
  const busy = ref(false);
  const scope = effectScope();
  scopes.push(scope);
  const { undo, comments, sidePanel } = scope.run(() => {
    const stack = useSpreadsheetUndo({ model: wb, recalc: (_keys, done) => done([]), onApplied: () => triggerRef(wb), activateSheet: () => {} });
    const panel = useSheetSidePanel();
    return {
      undo: stack,
      sidePanel: panel,
      comments: useSheetComments({ book: { workbook: wb, activeSheet, activeIndex }, undo: stack, selection: { selection }, viewState, cell, sidePanel: panel, busy: () => busy.value }),
    };
  })!;
  const threadsOf = (index: number) => wb.value!.sheets[index]!.threads ?? [];
  return { wb, activeIndex, active, fields, cell, busy, undo, comments, sidePanel, threadsOf };
}

/** A keydown dispatched from inside a grid element. */
function gridKey(init: KeyboardEventInit): KeyboardEvent {
  const grid = document.createElement("div");
  grid.dataset.testid = "spreadsheet-grid";
  document.body.append(grid);
  const event = new KeyboardEvent("keydown", { bubbles: true, cancelable: true, ...init });
  grid.dispatchEvent(event);
  return event;
}

describe("useSheetComments — the panel's content (SPEC-135)", () => {
  it("groups threads by sheet in tab order, each sheet's by (row, col), with initials, times and own-message flags", () => {
    const { comments } = setup();
    const groups = comments.groups.value;
    expect(groups.map((group) => group.sheetName)).toEqual(["Summary", "Data"]);
    const summary = groups[0]!.items;
    expect(summary.map((item) => (item.kind === "thread" ? item.address : ""))).toEqual(["D2", "E2"]);
    const d2 = summary[0]!;
    expect(d2.kind === "thread" && d2.messages.map((m) => [m.author, m.initials, m.own])).toEqual([
      ["Sam Lee", "SL", false],
      ["Alex Kim", "AK", false],
      ["You", "Y", true],
    ]);
    expect(d2.kind === "thread" && d2.messages[0]!.time).toBe(messageTime("2026-03-01T09:00:00Z"));
    expect(messageTime(undefined)).toBe("");
    expect(messageTime("not a date")).toBe("");
  });

  it("the unresolved filter hides resolved threads; an unknown filter value falls back to all", () => {
    const { comments } = setup();
    comments.setFilter("unresolved");
    expect(comments.groups.value.flatMap((group) => group.items.map((item) => item.kind === "thread" && item.threadId))).toEqual(["{D2}", "{A1}"]);
    comments.setFilter("anything");
    expect(comments.filter.value).toBe("all");
  });

  it("in an .ods workbook lists the notes instead, and New comment is off with ODS_NOTES_ONLY_TITLE", () => {
    const { comments } = setup("ods");
    expect(comments.notesOnly.value).toBe(true);
    expect(comments.groups.value).toEqual([
      { sheetId: "s1", sheetName: "Summary", items: [expect.objectContaining({ kind: "note", noteId: "n-B2", address: "B2", author: "Alex Kim", initials: "AK", text: "Q3 figure, unaudited." })] },
    ]);
    expect(comments.canComment.value).toBe(false);
    expect(comments.newCommentDisabledReason.value).toBe(ODS_NOTES_ONLY_TITLE);
    expect(comments.newComment()).toBe(false);
    expect(comments.composer.value).toBeNull();
  });

  it("the identity field saves the name (a blank one resets to You), and own-message flags follow it", () => {
    const { comments } = setup();
    comments.setAuthor("Sam Lee");
    expect(comments.author.value).toBe("Sam Lee");
    const d2 = comments.groups.value[0]!.items[0]!;
    expect(d2.kind === "thread" && d2.messages.map((m) => m.own)).toEqual([true, false, false]);
    comments.setAuthor("   ");
    expect(comments.author.value).toBe("You");
  });
});

describe("useSheetComments — the commands (SPEC-135)", () => {
  it("New comment opens the composer on the active cell and the pane; submitting adds a thread as 'Add comment'", () => {
    const { comments, active, undo, threadsOf, sidePanel } = setup();
    active.value = { row: 6, col: 2 };
    expect(comments.newComment()).toBe(true);
    expect(sidePanel.activePane.value).toBe("comments");
    expect(comments.composer.value).toEqual({ sheetId: "s1", ref: { row: 6, col: 2 }, address: "C7" });
    comments.composerDraft.value = "Is this right?";
    comments.submitComposer();
    expect(comments.composer.value).toBeNull();
    const added = threadsOf(0).find((each) => each.ref.row === 6)!;
    expect(added.messages[0]).toMatchObject({ author: "You", text: "Is this right?" });
    expect(comments.focusedId.value).toBe(added.threadId);
    expect(undo.undoLabel.value).toBe(COMMENT_UNDO_LABELS.add);
    undo.undo();
    expect(threadsOf(0).some((each) => each.ref.row === 6)).toBe(false);
  });

  it("New comment on a cell with a thread opens the pane on that thread; Ctrl+Alt+M from the grid runs it; a busy grid does nothing", () => {
    const { comments, active, busy } = setup();
    active.value = { row: 1, col: 3 };
    const event = gridKey({ key: "m", code: "KeyM", ctrlKey: true, altKey: true, metaKey: false });
    // Cmd+Option+M on macOS, Ctrl+Alt+M elsewhere: whichever the platform reads.
    if (!event.defaultPrevented) {
      gridKey({ key: "µ", code: "KeyM", metaKey: true, altKey: true });
    }
    expect(comments.composer.value).toBeNull();
    expect(comments.focusedId.value).toBe("{D2}");
    busy.value = true;
    comments.focusedId.value = null;
    expect(comments.newComment()).toBe(false);
    expect(comments.focusedId.value).toBeNull();
  });

  it("reply, resolve and reopen are each one labelled step", () => {
    const { comments, undo, threadsOf } = setup();
    comments.setReplyDraft("{D2}", "One more question.");
    comments.reply("{D2}");
    expect(threadsOf(0)[0]!.messages.at(-1)).toMatchObject({ author: "You", text: "One more question." });
    expect(comments.replyDrafts.value["{D2}"]).toBeUndefined();
    expect(undo.undoLabel.value).toBe(COMMENT_UNDO_LABELS.reply);
    comments.resolve("{D2}");
    expect(threadsOf(0)[0]!.resolved).toBe(true);
    expect(undo.undoLabel.value).toBe(COMMENT_UNDO_LABELS.resolve);
    comments.reopen("{D2}");
    expect(threadsOf(0)[0]!.resolved).toBe(false);
    expect(undo.undoLabel.value).toBe(COMMENT_UNDO_LABELS.reopen);
    undo.undo();
    expect(threadsOf(0)[0]!.resolved).toBe(true);
    // A blank reply is nothing.
    comments.setReplyDraft("{D2}", "  ");
    comments.reply("{D2}");
    expect(undo.undoLabel.value).toBe(COMMENT_UNDO_LABELS.resolve);
  });

  it("edits and deletes only the identity's own messages; deleting the opening message deletes the thread", () => {
    const { comments, undo, threadsOf } = setup();
    comments.startEdit("{D2}", "{D2}");
    expect(comments.editing.value).toBeNull();
    comments.deleteMessage("{D2}", "{D2}-1");
    expect(threadsOf(0)[0]!.messages).toHaveLength(3);
    comments.startEdit("{D2}", "{D2}-2");
    expect(comments.editDraft.value).toBe("Thanks.");
    comments.editDraft.value = "Thanks, that helps.";
    comments.saveEdit();
    expect(threadsOf(0)[0]!.messages[2]!.text).toBe("Thanks, that helps.");
    expect(undo.undoLabel.value).toBe(COMMENT_UNDO_LABELS.edit);
    comments.deleteMessage("{D2}", "{D2}-2");
    expect(threadsOf(0)[0]!.messages).toHaveLength(2);
    expect(undo.undoLabel.value).toBe(COMMENT_UNDO_LABELS.delete);
    comments.deleteMessage("{A1}", "{A1}");
    expect(threadsOf(1)).toEqual([]);
  });
});

describe("useSheetComments — the pane and navigation (SPEC-135)", () => {
  it("toggle opens and closes the comments pane; openThread opens it on a thread, showing a resolved one under the filter", () => {
    const { comments, sidePanel } = setup();
    comments.toggle();
    expect(comments.isOpen.value).toBe(true);
    comments.toggle();
    expect(sidePanel.activePane.value).toBeNull();
    comments.setFilter("unresolved");
    comments.openThread("s1", "{E2}");
    expect(comments.isOpen.value).toBe(true);
    expect(comments.focusedId.value).toBe("{E2}");
    expect(comments.filter.value).toBe("all");
  });

  it("activating a thread selects its sheet and cell and reveals it", async () => {
    const { comments, cell, activeIndex, fields } = setup();
    comments.activate("s2", "{A1}", { row: 0, col: 0 });
    expect(cell.selectSheet).toHaveBeenCalledWith(1);
    expect(activeIndex.value).toBe(1);
    expect(fields.value.s2?.selection?.active).toEqual({ row: 0, col: 0 });
    await nextTick();
    expect(cell.revealTick.value).toBe(1);
    comments.activate("s1", "{D2}", { row: 1, col: 3 });
    expect(fields.value.s1?.selection?.active).toEqual({ row: 1, col: 3 });
  });
});
