// The document outline navigator (SPEC-037): the pane's persisted open
// state and collapse sets, the outline tree with its active-heading
// highlight, the navigate/collapse/move handlers, and the body heading
// chevron (AC-16) whose hidden-block set pagination lays out around. Two
// pieces stay in DocumentView: the `caretPosition` mirror this reads
// (updateCaretStatus writes it with the rest of the caret tracking), and
// the watch(hiddenBodyBlocks) that re-schedules pagination — the controller
// is created after this composable is wired, precisely because its layout
// closure reads the hiddenBodyBlocks returned here.
import { computed, ref, watch } from "vue";
import {
  activeHeadingAnchor,
  blockKey,
  buildOutline,
  emptyKeySet,
  hiddenBlockAddresses,
  moveHeading,
} from "../../lib/outline";
import { loadOutlinePaneOpen, saveOutlinePaneOpen } from "../../lib/outlinePane";
import type { CaretPosition } from "../../lib/textEditing";
import type { EditorContext } from "./context";

export interface OutlinePaneDeps extends EditorContext {
  /** Live read of the view's reactive caret mirror (`caretPosition`) — a
   * getter over a ref, so activeOutlineAnchor re-runs when the caret
   * moves. */
  getCaretPosition: () => CaretPosition | null;
  /** useLinks' anchor jump (the same jump the TOC and internal links use).
   * useLinks is wired later in the view than this composable, so it comes
   * in as a deferring lambda read at call time. */
  jumpToAnchor: (anchor?: string) => void;
}

export function useOutlinePane(deps: OutlinePaneDeps) {
  const {
    doc,
    contentRoot,
    editable,
    currentSelection,
    applyEdit,
    applyCaretMove,
    getCaret,
    getCaretPosition,
    jumpToAnchor,
  } = deps;

  // Per-user open/closed state, persisted to localStorage (AC-14) and open by
  // default for a new user (AC-15).
  const outlinePaneOpen = ref(loadOutlinePaneOpen());
  watch(outlinePaneOpen, (open) => saveOutlinePaneOpen(open));
  // Transient view state, never part of the model: which pane tree nodes are
  // collapsed (AC-7, AC-8), and which body headings are collapsed in the
  // document body (AC-16), keyed by "section:block". Body collapse hides blocks
  // from layout only — it is never written to the file (AC-17).
  const outlinePaneCollapsed = ref(emptyKeySet());
  const collapsedHeadings = ref(emptyKeySet());

  const outlineTree = computed(() => (doc.value ? buildOutline(doc.value) : []));
  const activeOutlineAnchor = computed(() =>
    doc.value ? activeHeadingAnchor(doc.value, getCaretPosition()) : null,
  );
  const hiddenBodyBlocks = computed(() =>
    doc.value ? hiddenBlockAddresses(doc.value, collapsedHeadings.value) : emptyKeySet(),
  );

  // --- outline navigator handlers (SPEC-037) ---------------------------------
  /** View menu toggle / persistence of the pane's open state (AC-13, AC-14). */
  function toggleOutlinePane() {
    outlinePaneOpen.value = !outlinePaneOpen.value;
  }

  /** Clicking or Enter on an outline entry (AC-5, AC-6, AC-19): jump to the
   * heading — place the caret at its start and scroll it into view. Reuses the
   * same anchor-jump the TOC and internal links use. */
  function handleOutlineNavigate(anchor = "") {
    jumpToAnchor(anchor);
  }

  /** Collapse/expand a node in the pane (AC-7, AC-8) — pane-only view state. */
  function handleOutlineToggleCollapse(anchor = "") {
    const next = new Set(outlinePaneCollapsed.value);
    if (next.has(anchor)) {
      next.delete(anchor);
    } else {
      next.add(anchor);
    }
    outlinePaneCollapsed.value = next;
  }

  /** Drop of a dragged outline entry (AC-9, AC-10, AC-11): move the heading and
   * all its content as ONE undo step. moveHeading returns null for a rejected
   * drop (different level, non-sibling, onto itself), so no edit is recorded. */
  function handleOutlineMove(payload = { source: "", target: "", position: "before" }) {
    const model = doc.value;
    if (!model || !editable.value) {
      return;
    }
    const before = currentSelection(contentRoot.value as HTMLElement | null);
    // Narrow the payload's position to the DropPosition literal union (espree
    // bars a type annotation on the handler's own parameter).
    const position = payload.position === "after" ? "after" : "before";
    const result = moveHeading(model, payload.source, payload.target, position);
    if (result) {
      applyEdit(result, before, "move heading");
    }
  }

  /** The body chevron next to a heading (AC-16): collapse/expand its content.
   * Keyed by "section:block"; purely transient (AC-17). When collapsing the
   * section the caret sits in, move the caret to the heading so it never hides
   * behind collapsed content. */
  function handleBodyHeadingToggle(section = 0, block = 0) {
    const model = doc.value;
    if (!model) {
      return;
    }
    const key = blockKey(section, block);
    const next = new Set(collapsedHeadings.value);
    if (next.has(key)) {
      next.delete(key);
    } else {
      next.add(key);
    }
    collapsedHeadings.value = next;
    const hidden = hiddenBlockAddresses(model, next);
    const caret = getCaret();
    if (caret && hidden.has(blockKey(caret.section, caret.block))) {
      applyCaretMove({ section, block, offset: 0 });
    }
  }

  return {
    outlinePaneOpen,
    outlinePaneCollapsed,
    collapsedHeadings,
    outlineTree,
    activeOutlineAnchor,
    hiddenBodyBlocks,
    toggleOutlinePane,
    handleOutlineNavigate,
    handleOutlineToggleCollapse,
    handleOutlineMove,
    handleBodyHeadingToggle,
  };
}
