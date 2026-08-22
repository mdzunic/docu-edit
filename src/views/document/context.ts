// The shared dependency contract between DocumentView (the editor backbone)
// and the feature composables in this directory.
//
// Before this file existed, every composable declared its own `*Deps`
// interface and re-spelled the same core fields — `doc` was declared 22
// times, `applyEdit` 15 times in five drifting shapes (several re-inlining
// the literal definition of CommandResult), and refs whose real types exist
// in documentModel.ts were typed `Ref<any>` because each file re-derived its
// own contract. One declaration here ends the drift: composables extend
// EditorContext with only their genuinely feature-specific fields, and the
// view builds a single context object it spreads into every `use*` call.
//
// Field types follow what the view actually provides:
// - `doc` is a ShallowRef on purpose — the model must stay a plain object
//   (IndexedDB structured clone rejects reactive proxies) and re-renders via
//   triggerRef; see the `doc` declaration in DocumentView.vue.
// - `contentRoot`/`pageRef` are `Element`, not `HTMLElement`: the view's
//   template refs default to that type, and its `.vue` script cannot cast
//   (ESLint parses `.vue` scripts with espree, which bars TS syntax), so
//   composables cast `.value` where they need the narrower type.
// - Getters (`getCaret`, `getScrollContainer`) wrap plain variables that are
//   deliberately not refs — nothing renders from them, and they must be read
//   live, never destructured into a snapshot.
import type { Ref, ShallowRef } from "vue";
import type { CommandResult, ModelSelection } from "../../lib/commands";
import type { ActiveHeaderFooterTarget, DocumentModel } from "../../lib/documentModel";
import type { CaretPosition } from "../../lib/textEditing";

/** The optional behaviours of an applyEdit call. The command result and the
 * before-selection are data every edit has; these are the exceptions a call
 * site opts into — as named flags, so a call reads as
 * `applyEdit(result, before, "paste", { keepPasteChip: true })` rather than
 * a tail of positional booleans. */
export interface ApplyEditOptions {
  /** Consecutive groupable edits with the same label fold into one undo
   * step (SPEC-015 AC-9: literal typing groups; everything else doesn't). */
  groupable?: boolean;
  /** Keeps the paste-options chip alive across this edit (SPEC-016): the
   * paste itself and the chip's own re-application set it; any other edit
   * supersedes the chip. */
  keepPasteChip?: boolean;
  /** Skips the track-changes rewrite (SPEC-039): for the accept/reject
   * commands themselves, whose result IS the reviewer's chosen final state,
   * not a new edit to re-diff. */
  skipTrackChanges?: boolean;
  /** Skips placing the DOM caret/selection back in the canvas after this
   * edit (SPEC-056): the equation insert paths hand focus straight to the
   * equation editor's own surface, and restoreCaretAfterRender's DOM-selection
   * write would otherwise implicitly refocus the contenteditable root and
   * steal the keystrokes the user means for the equation. */
  skipCaretRestore?: boolean;
  /** Skips the keepCaretVisible() scroll nudge after this edit (SPEC-064): a
   * vertical-ruler drag fires one edit per pointermove while the user has
   * scrolled the caret off-screen to reach a bottom-margin/footer/row marker.
   * Nudging the caret back into view mid-drag would jump the viewport and
   * desync the fixed on-screen pointer from clientToContentPx's page-rect
   * mapping, so every following pointermove reads a wildly wrong value. The
   * caret has not moved — the user chose this scroll position — so leave it. */
  skipCaretScroll?: boolean;
}

/** Typed empty-options seed for the `.vue` script's applyEdit default
 * parameter (espree bars an annotation there — same idiom as the `no*`
 * seeds in lib/). */
export function noApplyEditOptions(): ApplyEditOptions {
  return {};
}

/** What the DocumentView backbone provides to every feature composable.
 *
 * A composable's `*Deps` interface extends this and adds only its own
 * fields; the view builds one object satisfying it and spreads that into
 * each `use*({ ...editorContext, ...extras })` call. A composable is free
 * to ignore fields it does not need — uniformity of the contract is the
 * point, so wiring a new feature never means re-plumbing the core. */
export interface EditorContext {
  /** The open document model, replaced (not mutated) on every edit. */
  doc: ShallowRef<DocumentModel | null>;
  documentId: Readonly<Ref<string>>;
  /** The contenteditable canvas root (`.document-content`). */
  contentRoot: Ref<Element | null>;
  /** The `.pages` container — overlay rects are positioned relative to it. */
  pageRef: Ref<Element | null>;
  /** Whether the canvas accepts edits (lock owned, not previewing history,
   * under the supported ceiling). Read commands gate on their own broader
   * conditions — see findAndPrintEnabled in the view. */
  editable: Readonly<Ref<boolean>>;
  /** The current selection in model coordinates: the live DOM selection
   * when it is inside the canvas, else the tracked selection/caret. */
  currentSelection: (root?: HTMLElement | null) => ModelSelection | null;
  /** Applies a command result as the new document state and records it in
   * the undo history under `label`. See ApplyEditOptions for the flags. */
  applyEdit: (
    result: CommandResult | null,
    beforeSelection: ModelSelection | null,
    label: string,
    options?: ApplyEditOptions,
  ) => void;
  /** Moves the caret (or, with extend, the selection's focus end) to a
   * model position and refreshes tracking + announcement. */
  applyCaretMove: (target?: CaretPosition, extend?: boolean) => void;
  /** Brings a model position into view, framed a third of the way down. */
  scrollModelCaretIntoView: (position?: CaretPosition, always?: boolean) => Promise<void>;
  /** Live access to the view's raw caret — read it fresh at call time. */
  getCaret: () => CaretPosition | null;
  /** The scrollable `<main>` ancestor of the canvas, once mounted. */
  getScrollContainer: () => Element | null;
  /** Clears the table cell-range selection (a sibling overlay). */
  clearCellSelection: () => void;
  addToast: (message: string) => void;
  /** Marks the session dirty: feeds autosave and the version-history
   * cadence timer (both hang off every real edit). */
  markDirty: () => void;
  /** Mirrors the undo history's per-document state into the toolbar. */
  refreshUndoRedoState: () => void;
  /** Registers a callback run when the document session closes. */
  onSessionClose: (cb: () => void) => void;
  /** Registers a callback run on view unmount. */
  onUnmount: (cb: () => void) => void;
  /** Registers a callback run on canvas scroll (overlay re-measurement). */
  onScroll: (cb: () => void) => void;
  /** The presentation-layer zoom as a scale fraction (1 == 100 %),
   * SPEC-063. The model, pagination and stored geometry are unaffected —
   * this only scales the `.pages` transform and the coordinate helpers. */
  zoom: Readonly<Ref<number>>;
  /** Maps a client (viewport) point to unscaled content pixels relative to
   * the scaled `.pages` box — every pointer path measures through this so it
   * lands where the pointer is at any zoom (SPEC-063 AC-22, AC-24). */
  clientToContentPx: (clientX: number, clientY: number) => { x: number; y: number };
  /** The inverse: places a model-space content pixel into client
   * coordinates for a position: fixed overlay outside the transform subtree
   * (SPEC-063 AC-22, AC-57). */
  contentPxToClient: (contentX: number, contentY: number) => { x: number; y: number };
}

/** The header/footer editing bridge shared by composables whose features
 * can also act inside the active band (image and table insertion, SPEC-029
 * AC-5/AC-6): the active band, its selection, and its own commit path. */
export interface HeaderFooterBridge {
  activeHeaderFooter: Ref<ActiveHeaderFooterTarget | null>;
  currentHfSelection: () => ModelSelection | null;
  applyHeaderFooterEdit: (
    resultModel: DocumentModel,
    resultSelection: ModelSelection,
    label: string,
  ) => void;
}
