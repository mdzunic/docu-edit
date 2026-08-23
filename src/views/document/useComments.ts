// Comment threads (SPEC-038): the right-hand margin pane's state, the
// comments-only view's filters, the Insert -> Comment composer, and the
// highlight overlay + pane-card geometry. Comment editing itself is plain
// text (no synthetic rich-text model like useNotes/useHeadersFooters needs)
// — every mutation goes through the ordinary applyEdit, exactly like any
// other command.
import { computed, nextTick, ref, shallowRef, watch } from "vue";
import { domPointForCaret } from "../../lib/domCaret";
import { caretClientRect } from "../../lib/caretGeometry";
import { noIdleHandle, scheduleIdleWork, type IdleHandle } from "../../lib/core/idleScheduler";
import {
  addReply as addReplyCommand,
  canModifyMessage,
  commentThreadsAt,
  deleteMessage as deleteMessageCommand,
  deleteThread as deleteThreadCommand,
  editMessage as editMessageCommand,
  insertComment as insertCommentCommand,
  reopenThread as reopenThreadCommand,
  resolveThread as resolveThreadCommand,
  visibleCommentThreads,
  type VisibleCommentThread,
} from "../../lib/comments";
import { usePreferences } from "../../composables/usePreferences";
import { isCollapsed, type ModelSelection } from "../../lib/commands";
import type { CaretPosition } from "../../lib/textEditing";
import type { CommentMessage } from "../../lib/documentModel";
import type { EditorContext } from "./context";

/** One highlighted rectangle of an anchored (open, non-resolved) comment
 * range, page-relative like find's HighlightRect — plus which thread it
 * belongs to, so clicking it can focus that thread (AC "clicking a
 * highlighted range focuses its thread"). */
export interface CommentHighlightRect {
  left: number;
  top: number;
  width: number;
  height: number;
  threadId: string;
  active: boolean;
}

/** Typed empty seed for the highlight-rect array in `.vue` scripts (espree
 * bars the annotation — see other lib modules' noX() helpers). */
export function noCommentHighlightRects(): CommentHighlightRect[] {
  return [];
}

/** Typed pass-through for a CommentHighlightRect[] coming from a `.vue`
 * template's `unknown[]` prop (espree bars the annotation), so
 * PageMarkupOverlays can iterate the rects. */
export function asCommentHighlightRects(value: unknown): CommentHighlightRect[] {
  return value as CommentHighlightRect[];
}

export type CommentFilter = "all" | "open" | "resolved";

/** Typed pass-through for a comment filter coming from a `.vue` template's
 * plain string (same espree constraint as documentModel's
 * asCharacterFormatKey). */
export function asCommentFilter(value: string): CommentFilter {
  return value as CommentFilter;
}

/** Typed pass-through for a `VisibleCommentThread[]` arriving at
 * CommentsPane.vue as a runtime-declared Array prop (same espree constraint
 * as asCommentFilter — CommentsPane.vue cannot write the cast itself). */
export function asVisibleCommentThreads(value: unknown): VisibleCommentThread[] {
  return value as VisibleCommentThread[];
}

/** Typed pass-through for the threadId -> px `Map` arriving at
 * CommentsPane.vue as a runtime-declared Object prop. */
export function asCommentThreadTops(value: unknown): Map<string, number> {
  return value as Map<string, number>;
}

/** Typed pass-through for the author-name list arriving at CommentsPane.vue
 * as a runtime-declared Array prop. */
export function asCommentAuthorList(value: unknown): string[] {
  return value as string[];
}

/** Typed pass-through for a `CommentMessage` arriving at CommentsPane.vue
 * from a `v-for` over an already-cast thread's messages. */
export function asCommentMessage(value: unknown): CommentMessage {
  return value as CommentMessage;
}

/** Typed `null` seed for the pane's single in-progress edit, in
 * CommentsPane.vue (same espree constraint as commands.ts's noSelection). */
export function noEditingMessageId(): string | null {
  return null;
}

/** Typed empty seed for the pane's per-thread reply drafts (a plain
 * Record, not a Map — `reactive()` needs a plain object to track key
 * additions). */
export function noReplyDrafts(): Record<string, string> {
  return {};
}

export type CommentsDeps = EditorContext;

export function useComments(deps: CommentsDeps) {
  const {
    doc,
    contentRoot,
    pageRef,
    editable,
    currentSelection,
    applyEdit,
    applyCaretMove,
    scrollModelCaretIntoView,
    getCaret,
    addToast,
    onSessionClose,
    onUnmount,
    onScroll,
    zoom,
  } = deps;

  // --- pane / view state (per-user, not part of the model) -----------------

  // Open/closed state. Since SPEC-065 this is a tab in the consolidated side
  // panel, which owns the persistence (one key, migrated from the legacy
  // docu-edit:comments-pane key) and drives this ref through useSidePanel.
  const commentsPaneOpen = ref(false);
  function toggleCommentsPane(): void {
    commentsPaneOpen.value = !commentsPaneOpen.value;
  }

  const commentsOnlyView = ref(false);
  function toggleCommentsOnlyView(): void {
    commentsOnlyView.value = !commentsOnlyView.value;
  }

  const commentFilter = ref<CommentFilter>("all");
  const commentAuthorFilter = ref<string>("");

  // The author identity is the SPEC-038 settable name, now one of the
  // centralized preferences (SPEC-066) rather than its own docu-edit:comment-
  // author key. A writable computed so the comments UI still reads and sets it
  // as a ref, but every write goes through the store (AC-18, AC-46).
  const preferencesApi = usePreferences();
  const commentAuthorName = computed({
    get: () => preferencesApi.preferences.authorName,
    set: (name: string) => preferencesApi.setPreference("authorName", name),
  });

  const activeThreadId = ref<string | null>(null);

  // --- derived thread lists --------------------------------------------------

  const visibleThreads = computed<VisibleCommentThread[]>(() =>
    doc.value ? visibleCommentThreads(doc.value) : [],
  );
  const openThreads = computed(() => visibleThreads.value.filter((entry) => !entry.thread.resolved));
  const resolvedThreads = computed(() => visibleThreads.value.filter((entry) => entry.thread.resolved));

  /** Every distinct message author across every visible thread, in first-
   * appearance (document) order — the comments-only view's author filter
   * options (AC "filters... by author"). */
  const commentAuthors = computed<string[]>(() => {
    const seen = new Set<string>();
    const authors: string[] = [];
    for (const { thread } of visibleThreads.value) {
      for (const message of thread.messages) {
        if (!seen.has(message.author)) {
          seen.add(message.author);
          authors.push(message.author);
        }
      }
    }
    return authors;
  });

  /** The comments-only view's list (AC "lists all threads in document
   * order with filters for open/resolved and by author"). */
  const filteredThreads = computed(() =>
    visibleThreads.value.filter((entry) => {
      if (commentFilter.value === "open" && entry.thread.resolved) {
        return false;
      }
      if (commentFilter.value === "resolved" && !entry.thread.resolved) {
        return false;
      }
      if (
        commentAuthorFilter.value &&
        !entry.thread.messages.some((message) => message.author === commentAuthorFilter.value)
      ) {
        return false;
      }
      return true;
    }),
  );

  // --- composer (Insert -> Comment, AC-1..4) ---------------------------------

  const commentComposerOpen = ref(false);
  let commentComposerSelection: ModelSelection | null = null;

  /** "Insert -> Comment" (Ctrl/Cmd+Alt+M): opens the composer over the
   * current selection. A collapsed selection is a no-op — a comment always
   * anchors to a range of selected text (AC "Selecting text... highlights
   * the range"). */
  function openCommentComposer(): void {
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    if (!root || !model || !editable.value) {
      return;
    }
    const selection = currentSelection(root);
    if (!selection || isCollapsed(selection)) {
      addToast("Select some text first to add a comment");
      return;
    }
    commentComposerSelection = selection;
    commentsPaneOpen.value = true;
    commentComposerOpen.value = true;
  }

  function cancelCommentComposer(): void {
    commentComposerOpen.value = false;
    commentComposerSelection = null;
  }

  /** Submits the composer's text as a new thread (AC "Submitting adds the
   * comment with an author name and timestamp"). Blank text is a no-op. */
  function submitComment(text = ""): void {
    const model = doc.value;
    const selection = commentComposerSelection;
    const trimmed = text.trim();
    if (!model || !selection || trimmed.length === 0) {
      return;
    }
    const result = insertCommentCommand(model, selection, commentAuthorName.value, trimmed);
    applyEdit(result, selection, "insert comment");
    commentComposerOpen.value = false;
    commentComposerSelection = null;
  }

  // --- thread mutations -------------------------------------------------------

  function currentModelSelection(): ModelSelection {
    const at = getCaret();
    return at ? { anchor: at, focus: at } : { anchor: { section: 0, block: 0, offset: 0 }, focus: { section: 0, block: 0, offset: 0 } };
  }

  function submitReply(threadId: string, text: string): void {
    const model = doc.value;
    const trimmed = text.trim();
    if (!model || trimmed.length === 0) {
      return;
    }
    const selection = currentModelSelection();
    applyEdit(
      addReplyCommand(model, selection, threadId, commentAuthorName.value, trimmed),
      selection,
      "reply to comment",
    );
  }

  /** Whether the current identity may edit/delete `message` (AC "author can
   * edit/delete their own comments"). */
  function canModifyOwnMessage(message: CommentMessage): boolean {
    return canModifyMessage(message, commentAuthorName.value);
  }

  function saveEditedMessage(threadId: string, messageId: string, text: string): void {
    const model = doc.value;
    const trimmed = text.trim();
    if (!model || trimmed.length === 0) {
      return;
    }
    const selection = currentModelSelection();
    applyEdit(
      editMessageCommand(model, selection, threadId, messageId, trimmed),
      selection,
      "edit comment",
    );
  }

  function deleteMessage(threadId: string, messageId: string): void {
    const model = doc.value;
    if (!model) {
      return;
    }
    const selection = currentModelSelection();
    applyEdit(deleteMessageCommand(model, selection, threadId, messageId), selection, "delete comment");
  }

  function resolveThread(threadId: string): void {
    const model = doc.value;
    if (!model) {
      return;
    }
    const selection = currentModelSelection();
    applyEdit(resolveThreadCommand(model, selection, threadId), selection, "resolve comment");
  }

  function reopenThread(threadId: string): void {
    const model = doc.value;
    if (!model) {
      return;
    }
    const selection = currentModelSelection();
    applyEdit(reopenThreadCommand(model, selection, threadId), selection, "reopen comment");
  }

  /** Deletes a whole thread (AC "deleting a thread removes its anchor") —
   * covered by the ordinary undo stack like every applyEdit call, so no
   * bespoke undo affordance is needed here (see the AC-17 toast below for
   * the *implicit* deletion case, where there is no explicit user action to
   * hang a confirmation off). */
  function deleteThread(threadId: string): void {
    const model = doc.value;
    if (!model) {
      return;
    }
    const selection = currentModelSelection();
    if (activeThreadId.value === threadId) {
      activeThreadId.value = null;
    }
    applyEdit(deleteThreadCommand(model, selection, threadId), selection, "delete comment thread");
  }

  // --- focus / navigation (AC-7, AC-8) ----------------------------------------

  /** Clicking a highlighted range in the body focuses its thread in the
   * pane (AC-7) — no body scroll needed, the click already happened there. */
  function focusThreadFromHighlight(threadId: string): void {
    activeThreadId.value = threadId;
    commentsPaneOpen.value = true;
  }

  /** Clicking a thread (in the pane, or the comments-only list) scrolls to
   * and highlights its anchored range (AC-8). */
  function focusThread(threadId: string): void {
    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    if (!model) {
      return;
    }
    activeThreadId.value = threadId;
    const entry = visibleCommentThreads(model).find((item) => item.thread.threadId === threadId);
    if (!entry) {
      return;
    }
    root?.focus();
    applyCaretMove(entry.start);
    void scrollModelCaretIntoView(entry.start, true);
  }

  /** The thread(s) the caret currently sits inside — drives a "you're
   * inside a comment" indicator if the caller wants one. */
  function threadsAtCaret(caret: CaretPosition): VisibleCommentThread[] {
    const model = doc.value;
    return model ? commentThreadsAt(model, caret) : [];
  }

  // --- highlight overlay + pane geometry --------------------------------------

  const commentHighlightRects = shallowRef<CommentHighlightRect[]>(noCommentHighlightRects());
  /** Each open thread's card `top` (px), relative to the comments dock
   * container passed to measureCommentPositions — the margin-pane
   * counterpart of commentHighlightRects. */
  const commentThreadTops = shallowRef<Map<string, number>>(new Map());

  function rangeClientRects(root: HTMLElement, start: CaretPosition, end: CaretPosition): DOMRect[] {
    const startPoint = domPointForCaret(root, start);
    const endPoint = domPointForCaret(root, end);
    if (!startPoint || !endPoint) {
      return [];
    }
    const range = root.ownerDocument.createRange();
    try {
      range.setStart(startPoint.node, startPoint.offset);
      range.setEnd(endPoint.node, endPoint.offset);
    } catch {
      return [];
    }
    return Array.from(range.getClientRects());
  }

  /** Rebuilds the highlight overlay (open threads only — AC "resolving...
   * removes the highlight") and the pane's per-thread card positions. Runs
   * off the keystroke/scroll path via idle scheduling, mirroring
   * useFindReplace's measureFindHighlights. */
  function measureComments(dockEl?: Element | null): void {
    const root = contentRoot.value as HTMLElement | null;
    const pagesEl = pageRef.value;
    if (!root || !pagesEl) {
      commentHighlightRects.value = noCommentHighlightRects();
      return;
    }
    const pagesRect = pagesEl.getBoundingClientRect();
    // The highlight layer paints INSIDE the scaled `.pages` container, so the
    // rects carry unscaled, .pages-relative px — the CSS transform re-applies
    // the zoom, keeping the highlight registered with its text (SPEC-063
    // AC-25). Dividing by the scale undoes what getBoundingClientRect folded in.
    const z = zoom.value || 1;
    const rects: CommentHighlightRect[] = [];
    for (const { thread, start, end } of openThreads.value) {
      const active = activeThreadId.value === thread.threadId;
      for (const rect of rangeClientRects(root, start, end)) {
        rects.push({
          left: (rect.left - pagesRect.left) / z,
          top: (rect.top - pagesRect.top) / z,
          width: rect.width / z,
          height: rect.height / z,
          threadId: thread.threadId,
          active,
        });
      }
    }
    commentHighlightRects.value = rects;

    if (dockEl) {
      const dockRect = dockEl.getBoundingClientRect();
      const tops = new Map<string, number>();
      for (const { thread, start } of visibleThreads.value) {
        const caretRect = caretClientRect(root, start);
        if (caretRect) {
          tops.set(thread.threadId, caretRect.top - dockRect.top);
        }
      }
      commentThreadTops.value = tops;
    }
  }

  let commentsIdleHandle: IdleHandle | null = noIdleHandle();
  let lastDockEl: Element | null = null;
  function scheduleCommentsMeasure(dockEl?: Element | null): void {
    if (dockEl !== undefined) {
      lastDockEl = dockEl;
    }
    if (commentsIdleHandle) {
      return;
    }
    commentsIdleHandle = scheduleIdleWork(
      () => {
        commentsIdleHandle = null;
        measureComments(lastDockEl);
      },
      { timeoutMs: 150, fallbackBudgetMs: 8, maxBudgetMs: 16 },
    );
  }

  function cancelCommentsScheduled(): void {
    if (commentsIdleHandle) {
      commentsIdleHandle.cancel();
      commentsIdleHandle = null;
    }
  }

  watch(doc, () => void nextTick().then(() => scheduleCommentsMeasure()));
  watch(activeThreadId, () => void nextTick().then(() => scheduleCommentsMeasure()));
  onScroll(() => scheduleCommentsMeasure());
  onSessionClose(() => {
    cancelCommentsScheduled();
    commentHighlightRects.value = noCommentHighlightRects();
    commentThreadTops.value = new Map();
    activeThreadId.value = null;
    commentComposerOpen.value = false;
    commentComposerSelection = null;
  });
  onUnmount(() => cancelCommentsScheduled());

  // --- AC "deleting the entire range deletes the comment with an
  // undoable warning" -----------------------------------------------------
  // There is no bespoke "undo this" toast action anywhere else in the app
  // (see useToasts.ts) — every applyEdit-routed change, this one included,
  // is already on the ordinary Ctrl/Cmd+Z stack. So "undoable" is satisfied
  // structurally; this watcher only adds the discoverability nudge: when an
  // edit makes a thread that was visible a moment ago stop being visible,
  // while its record is STILL in model.comments (an explicit deleteThread
  // call removes both at once, so this never fires for that case — see
  // lib/comments.ts's header comment on derived visibility), tell the user
  // it happened and that it can be undone.
  let knownVisibleThreadIds = new Set<string>();
  watch(
    doc,
    (model) => {
      if (!model) {
        knownVisibleThreadIds = new Set();
        return;
      }
      const nowVisible = new Set(visibleCommentThreads(model).map((entry) => entry.thread.threadId));
      for (const threadId of knownVisibleThreadIds) {
        if (!nowVisible.has(threadId) && model.comments?.some((thread) => thread.threadId === threadId)) {
          addToast("Comment deleted — Ctrl+Z / Cmd+Z to undo");
        }
      }
      knownVisibleThreadIds = nowVisible;
    },
    { immediate: true },
  );

  return {
    commentsPaneOpen,
    toggleCommentsPane,
    commentsOnlyView,
    toggleCommentsOnlyView,
    commentFilter,
    commentAuthorFilter,
    commentAuthorName,
    activeThreadId,
    visibleThreads,
    openThreads,
    resolvedThreads,
    commentAuthors,
    filteredThreads,
    commentComposerOpen,
    openCommentComposer,
    cancelCommentComposer,
    submitComment,
    submitReply,
    canModifyOwnMessage,
    saveEditedMessage,
    deleteMessage,
    resolveThread,
    reopenThread,
    deleteThread,
    focusThreadFromHighlight,
    focusThread,
    threadsAtCaret,
    commentHighlightRects,
    commentThreadTops,
    scheduleCommentsMeasure,
  };
}
