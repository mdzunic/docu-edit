// Comments on cells — Excel 365's threads — and the comments pane (SPEC-135
// "Threads"). The model and its commands are lib/spreadsheet/annotations.ts';
// this composable drives SheetCommentsPanel (`sheet-comments-panel`), the
// `"comments"` pane of SPEC-129's sheet side panel:
// - View ▸ Comments and the app bar's `sheet-comments-toggle` toggle the pane
//   (`toggle`); `openThread` opens it on one thread (the popover's
//   `sheet-thread-open`); activating a thread selects its sheet and cell.
// - `groups`: the workbook's threads by sheet in tab order, each sheet's by
//   (row, col), under the filter (`all` | `unresolved`); in an .ods workbook,
//   which holds notes only, the workbook's notes instead (`notesOnly`).
// - New comment — Ctrl+Alt+M / Cmd+Option+M (`sheet.newComment`, from the
//   grid only; best effort on macOS, whose Option+Cmd+M minimises the
//   windows), Insert ▸ Comment and the cell context menu — opens the composer
//   (`sheet-comment-composer`) on the active cell, or the cell's thread when
//   it has one. In an .ods workbook it is disabled with ODS_NOTES_ONLY_TITLE
//   (`newCommentDisabledReason`), on an .xlsx sheet whose source's comments
//   the reader could not model with NOTES_KEPT_AS_READ_TITLE (`blocked` —
//   AC-11: they stay verbatim, so a new one could not be saved); a CSV or
//   TSV workbook offers neither.
// - reply, resolve, reopen, edit and delete — the last two on the identity's
//   own messages only (canModifyMessage); deleting a thread's opening message
//   deletes the thread.
// - the identity (`author` — core/commentIdentity.ts, shared with the text
//   editor's comments).
// Every command is ONE SPEC-121 undo step on the thread's sheet, labelled
// COMMENT_UNDO_LABELS' ("Add comment", "Reply to comment", "Edit comment",
// "Resolve comment", "Reopen comment", "Delete comment"). Texts stay plain
// strings; the panel renders them with text interpolation only.
import { computed, nextTick, onScopeDispose, ref, shallowRef, type Ref, type ShallowRef } from "vue";
import { canModifyMessage, commentInitials } from "../../lib/core/comments";
import { loadCommentAuthor, saveCommentAuthor } from "../../lib/core/commentIdentity";
import { isMacPlatform } from "../../lib/core/platform";
import { formatTimestamp } from "../../lib/format";
import { findShortcut } from "../../lib/shortcuts";
import { formatA1, type CellRef } from "../../lib/spreadsheet/address";
import {
  addThread,
  COMMENT_UNDO_LABELS,
  deleteThreadMessage,
  editThreadMessage,
  ODS_NOTES_ONLY_TITLE,
  replyToThread,
  setThreadResolved,
  sheetNotes,
  sheetThreads,
  threadAt,
} from "../../lib/spreadsheet/annotations";
import { isCsvFormat, type Sheet, type SpreadsheetModel, type ThreadMessage } from "../../lib/spreadsheet/model";
import { singleCell } from "../../lib/spreadsheet/selection";
import { inGrid } from "./headerMenu";
import { SHEET_SHORTCUTS } from "./sheetShortcuts";
import type { CellEditing } from "./useCellEditing";
import type { SheetSelectionController } from "./useSheetSelection";
import type { SheetSidePanel } from "./useSheetSidePanel";
import type { SheetViewState } from "./useSheetViewState";
import type { SpreadsheetUndo } from "./useSpreadsheetUndo";

export type CommentFilter = "all" | "unresolved";

/** One message row as the panel shows it. */
export interface CommentMessageView {
  messageId: string;
  author: string;
  initials: string;
  time: string;
  text: string;
  /** The identity wrote it: it may be edited and deleted. */
  own: boolean;
}

/** One thread or (in an .ods workbook) one note in the panel. */
export type CommentItemView =
  | { kind: "thread"; threadId: string; ref: CellRef; address: string; resolved: boolean; messages: CommentMessageView[] }
  | { kind: "note"; noteId: string; ref: CellRef; address: string; author: string; initials: string; time: string; text: string };

/** A sheet's threads (or notes) under its name. */
export interface CommentGroupView {
  sheetId: string;
  sheetName: string;
  items: CommentItemView[];
}

export interface SheetCommentsDeps {
  book: {
    workbook: ShallowRef<SpreadsheetModel | null>;
    activeSheet: Readonly<Ref<Sheet | null>>;
    activeIndex: Ref<number>;
  };
  undo: Pick<SpreadsheetUndo, "record" | "commit">;
  /** The active cell New comment starts on. */
  selection: Pick<SheetSelectionController, "selection">;
  /** Activating a thread selects its cell (on its sheet). */
  viewState?: Pick<SheetViewState, "update">;
  /** The sheet switch (an edit in progress commits first) and the reveal. */
  cell?: Pick<CellEditing, "selectSheet" | "revealTick">;
  sidePanel: Pick<SheetSidePanel, "activePane" | "open" | "toggle">;
  /** Syncs the view state into the model before a step is recorded. */
  syncInto?: (model: SpreadsheetModel) => void;
  /** Whether a cell is being edited: New comment waits. */
  busy?: () => boolean;
  /** Why the sheet on show's comments cannot be changed ("" when they can) —
   * useAnnotationEditing's notesReason: its source broke the subset. */
  blocked?: () => string;
}

/** A message's time as the panel shows it ("" when it has none). */
export function messageTime(createdAt: string | undefined): string {
  const at = createdAt ? Date.parse(createdAt) : Number.NaN;
  return Number.isNaN(at) ? "" : formatTimestamp(new Date(at));
}

export function useSheetComments(deps: SheetCommentsDeps) {
  const model = (): SpreadsheetModel | null => deps.book.workbook.value;
  const busy = (): boolean => deps.busy?.() === true;

  const author = ref(loadCommentAuthor());
  /** The identity field: saved (a blank name resets to the default). */
  function setAuthor(name: string): void {
    saveCommentAuthor(name);
    author.value = loadCommentAuthor();
  }

  const filter = ref<CommentFilter>("all");
  function setFilter(value: string): void {
    filter.value = value === "unresolved" ? "unresolved" : "all";
  }

  /** An .ods workbook: notes only — the panel lists them, New comment is off. */
  const notesOnly = computed(() => model()?.format === "ods");
  /** Why New comment is unavailable ("" when it is available). */
  const newCommentDisabledReason = computed(() => (notesOnly.value ? ODS_NOTES_ONLY_TITLE : (deps.blocked?.() ?? "")));
  /** Whether New comment applies: an .xlsx grid sheet whose source's
   * comments were modelled. */
  const canComment = computed(() => {
    const m = model();
    return m !== null && !isCsvFormat(m.format) && !notesOnly.value && deps.book.activeSheet.value?.kind === "grid" && newCommentDisabledReason.value === "";
  });

  const isOpen = computed(() => deps.sidePanel.activePane.value === "comments");
  function toggle(): void {
    deps.sidePanel.toggle("comments");
  }

  /** The thread (or note) the panel points at — activated or opened on. */
  const focusedId = ref<string | null>(null);

  const messageView = (message: ThreadMessage): CommentMessageView => ({
    messageId: message.messageId,
    author: message.author,
    initials: commentInitials(message.author),
    time: messageTime(message.createdAt),
    text: message.text,
    own: canModifyMessage(message, author.value),
  });

  const groups = computed((): CommentGroupView[] => {
    const m = model();
    if (!m) {
      return [];
    }
    return m.sheets.flatMap((sheet): CommentGroupView[] => {
      const items: CommentItemView[] = notesOnly.value
        ? sheetNotes(sheet).map((note) => ({
            kind: "note",
            noteId: note.noteId,
            ref: note.ref,
            address: formatA1(note.ref),
            author: note.author,
            initials: commentInitials(note.author),
            time: messageTime(note.createdAt),
            text: note.text,
          }))
        : sheetThreads(sheet)
            .filter((thread) => filter.value === "all" || !thread.resolved)
            .map((thread) => ({
              kind: "thread",
              threadId: thread.threadId,
              ref: thread.ref,
              address: formatA1(thread.ref),
              resolved: thread.resolved,
              messages: thread.messages.map(messageView),
            }));
      return items.length > 0 ? [{ sheetId: sheet.id, sheetName: sheet.name, items }] : [];
    });
  });

  // --- the composer, replies, the message editor ---------------------------------------------------

  const composer = shallowRef<{ sheetId: string; ref: CellRef; address: string } | null>(null);
  const composerDraft = ref("");
  /** Reply drafts by thread id. */
  const replyDrafts = ref<Record<string, string>>({});
  const editing = shallowRef<{ threadId: string; messageId: string } | null>(null);
  const editDraft = ref("");

  /** One undo step on `sheetId`. */
  function run(label: string, sheetId: string, command: (m: SpreadsheetModel) => void): void {
    const m = model();
    if (!m) {
      return;
    }
    deps.syncInto?.(m);
    deps.undo.record(label, { sheetId });
    command(m);
    deps.undo.commit();
  }

  const sheetOfThread = (threadId: string): Sheet | undefined => model()?.sheets.find((sheet) => sheetThreads(sheet).some((thread) => thread.threadId === threadId));

  /** New comment on the active cell: the composer, or the cell's thread. */
  function newComment(): boolean {
    const sheet = deps.book.activeSheet.value;
    if (!sheet || !canComment.value || busy()) {
      return false;
    }
    const at = deps.selection.selection.value.active;
    const existing = threadAt(sheet, at);
    deps.sidePanel.open("comments");
    if (existing) {
      filter.value = existing.resolved ? "all" : filter.value;
      focusedId.value = existing.threadId;
      composer.value = null;
      return true;
    }
    composer.value = { sheetId: sheet.id, ref: { row: at.row, col: at.col }, address: formatA1(at) };
    composerDraft.value = "";
    focusedId.value = null;
    return true;
  }

  function submitComposer(): void {
    const open = composer.value;
    const text = composerDraft.value;
    if (!open || text.trim() === "") {
      return;
    }
    composer.value = null;
    composerDraft.value = "";
    let started: string | null = null;
    run(COMMENT_UNDO_LABELS.add, open.sheetId, (m) => {
      started = addThread(m, open.sheetId, open.ref, text, author.value)?.threadId ?? null;
    });
    focusedId.value = started;
  }

  function cancelComposer(): void {
    composer.value = null;
    composerDraft.value = "";
  }

  function setReplyDraft(threadId: string, text: string): void {
    replyDrafts.value = { ...replyDrafts.value, [threadId]: text };
  }

  function reply(threadId: string): void {
    const text = replyDrafts.value[threadId] ?? "";
    const sheet = sheetOfThread(threadId);
    if (!sheet || text.trim() === "") {
      return;
    }
    run(COMMENT_UNDO_LABELS.reply, sheet.id, (m) => void replyToThread(m, sheet.id, threadId, text, author.value));
    const { [threadId]: _sent, ...rest } = replyDrafts.value;
    replyDrafts.value = rest;
  }

  function setResolved(threadId: string, resolved: boolean): void {
    const sheet = sheetOfThread(threadId);
    if (sheet) {
      run(resolved ? COMMENT_UNDO_LABELS.resolve : COMMENT_UNDO_LABELS.reopen, sheet.id, (m) => void setThreadResolved(m, sheet.id, threadId, resolved));
    }
  }

  /** The message, when the identity may change it. */
  function ownMessage(threadId: string, messageId: string): { sheet: Sheet; message: ThreadMessage } | null {
    const sheet = sheetOfThread(threadId);
    const message = sheet && sheetThreads(sheet).find((thread) => thread.threadId === threadId)?.messages.find((each) => each.messageId === messageId);
    return sheet && message && canModifyMessage(message, author.value) ? { sheet, message } : null;
  }

  function startEdit(threadId: string, messageId: string): void {
    const own = ownMessage(threadId, messageId);
    if (own) {
      editing.value = { threadId, messageId };
      editDraft.value = own.message.text;
    }
  }

  function saveEdit(): void {
    const open = editing.value;
    const own = open ? ownMessage(open.threadId, open.messageId) : null;
    editing.value = null;
    if (!open || !own || editDraft.value === own.message.text) {
      return;
    }
    const text = editDraft.value;
    const label = text.trim() === "" ? COMMENT_UNDO_LABELS.delete : COMMENT_UNDO_LABELS.edit;
    run(label, own.sheet.id, (m) => void editThreadMessage(m, own.sheet.id, open.threadId, open.messageId, text));
  }

  function cancelEdit(): void {
    editing.value = null;
  }

  function deleteMessage(threadId: string, messageId: string): void {
    const own = ownMessage(threadId, messageId);
    if (own) {
      run(COMMENT_UNDO_LABELS.delete, own.sheet.id, (m) => void deleteThreadMessage(m, own.sheet.id, threadId, messageId));
    }
  }

  // --- navigation ----------------------------------------------------------------------------------

  /** Selects `ref` on `sheetId`, switching sheets as needed, and reveals it. */
  async function selectCellOn(sheetId: string, cell: CellRef): Promise<void> {
    const index = model()?.sheets.findIndex((sheet) => sheet.id === sheetId) ?? -1;
    if (index < 0) {
      return;
    }
    if (index !== deps.book.activeIndex.value) {
      deps.cell?.selectSheet(index);
    }
    deps.viewState?.update(sheetId, { selection: singleCell(cell) });
    await nextTick();
    if (deps.cell) {
      deps.cell.revealTick.value += 1;
    }
  }

  /** A thread (or an .ods note) activated in the panel: its sheet and cell selected. */
  function activate(sheetId: string, id: string, cell: CellRef): void {
    focusedId.value = id;
    void selectCellOn(sheetId, cell);
  }

  /** The panel opened on one thread (the popover's `sheet-thread-open`). */
  function openThread(_sheetId: string, threadId: string): void {
    deps.sidePanel.open("comments");
    const thread = sheetThreads(sheetOfThread(threadId) ?? { threads: [] }).find((each) => each.threadId === threadId);
    if (thread?.resolved) {
      filter.value = "all";
    }
    focusedId.value = threadId;
  }

  /** Ctrl+Alt+M / Cmd+Option+M from the grid: New comment. */
  function onKeydown(event: KeyboardEvent): void {
    if (event.defaultPrevented || !inGrid(event.target) || busy()) {
      return;
    }
    if (findShortcut(event, isMacPlatform(), SHEET_SHORTCUTS)?.id === "sheet.newComment") {
      event.preventDefault();
      newComment();
    }
  }
  if (typeof window !== "undefined") {
    window.addEventListener("keydown", onKeydown, true);
  }
  onScopeDispose(() => {
    if (typeof window !== "undefined") {
      window.removeEventListener("keydown", onKeydown, true);
    }
  });

  return {
    author,
    setAuthor,
    filter,
    setFilter,
    notesOnly,
    canComment,
    newCommentDisabledReason,
    isOpen,
    toggle,
    focusedId,
    groups,
    composer,
    composerDraft,
    newComment,
    submitComposer,
    cancelComposer,
    replyDrafts,
    setReplyDraft,
    reply,
    resolve: (threadId: string): void => setResolved(threadId, true),
    reopen: (threadId: string): void => setResolved(threadId, false),
    editing,
    editDraft,
    startEdit,
    saveEdit,
    cancelEdit,
    deleteMessage,
    activate,
    openThread,
    onKeydown,
  };
}

export type SheetComments = ReturnType<typeof useSheetComments>;

/** Narrows a component's runtime-typed `controller` prop (the .vue parser
 * bars casts). */
export function asSheetComments(value: unknown): SheetComments {
  return value as SheetComments;
}
