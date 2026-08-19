// Native drawing-group containers (SPEC-057): selection (click-to-enter,
// second-click-a-member, Escape-out), Group/Ungroup, the container and member
// gestures that reuse the pure operations in src/lib/drawingGroups.ts, the
// container/member toolbar actions, and keyboard proxy traversal. The pure
// model math lives in the lib module; this composable owns only the view
// state and the DOM measurement the layout-free lib deliberately does not.
// Runtime-safe TS (this is a .ts composable, not a .vue script) — see
// useShapes.ts, the sibling this mirrors.
import { computed, nextTick, ref } from "vue";
import {
  caretSelection,
  deleteBackward,
  deleteForward,
  deleteRange,
  documentStart,
  insertLineBreak,
  insertText,
  isCollapsed,
  noPendingFormat,
  noSelection,
  splitBlock,
  type ModelSelection,
} from "../../lib/commands";
import { CM_PER_PX } from "../../lib/core/units";
import { measureImageGeometry, noImageHandleGeometry } from "../../lib/caretGeometry";
import { asImageHandle, resizedImageSize } from "../../lib/imageResize";
import { unrotateDelta } from "../../lib/core/shapeGeometry";
import { modelRangeFromSelection, setDomCaret, setDomSelection } from "../../lib/domCaret";
import { blockLength, noCaret, type CaretPosition } from "../../lib/textEditing";
import { recordEdit as recordHistoryEdit } from "../../lib/undoHistory";
import { replaceDocumentModel } from "../../lib/documentStore";
import {
  isDrawingGroup,
  isSealedGroup,
  type DocumentModel,
  type DrawingGroup,
  type DrawingMember,
} from "../../lib/documentModel";
import {
  commitMemberTextModel,
  deleteGroupContainer,
  deleteMember,
  duplicateGroupContainer,
  duplicateMember,
  duplicatedMemberPath,
  ensureMemberTextFrame,
  groupDrawings,
  memberAtPath,
  memberShapeAt,
  moveGroupContainer,
  moveMember,
  nudgeGroupContainer,
  nudgeMember,
  noDrawingLocation,
  reorderMember,
  resizeGroupContainer,
  resizeMember,
  rotateGroupContainer,
  rotateMember,
  setGroupAltText,
  setGroupWrap,
  setMemberAltText,
  ungroupDrawing,
  withMemberTextModel,
  type DrawingLocation,
  type GroupEntry,
  type MemberReorder,
  type RectCm,
} from "../../lib/drawingGroups";
import { reanchorShape, reorderShapeZ, type ParagraphSpan, type ZOrderCommand } from "../../lib/shapes";
import { type EditorContext } from "./context";
import { asNudgeDirection, type NudgeDirection } from "../../lib/core/shapes";

/** What DocumentView passes in beyond the shared context: the ways to clear
 * the sibling per-kind selections (so a group selection is exclusive). */
export interface DrawingGroupsDeps extends EditorContext {
  clearOtherSelections: () => void;
  /** The current cross-kind multi-selection of top-level drawings (SPEC-057
   * "Selection across drawing kinds") — SPEC-051's shape Shift-click list,
   * widened to carry any drawing kind's `{ section, block }`. Group reads it
   * live; it is kind-agnostic in storage. */
  getGroupCandidates: () => { section: number; block: number }[];
}

/** A top-level drawing address in the unified multi-selection (SPEC-057
 * "Selection across drawing kinds") — shapes, images, charts and containers
 * all join one selection so Group spans kinds. */
interface DrawingRef {
  section: number;
  block: number;
}

export function useDrawingGroups(deps: DrawingGroupsDeps) {
  const { doc, contentRoot, pageRef, editable, documentId, currentSelection, applyEdit, getCaret, clearOtherSelections, getGroupCandidates, markDirty, refreshUndoRedoState } = deps;

  /** The selected container or member (SPEC-057): `path` empty = the
   * container itself, non-empty = a member by index path. */
  const groupSelection = ref<DrawingLocation | null>(noDrawingLocation());
  /** The unified multi-selection of top-level drawings, for Group. */
  const multiSelection = ref<DrawingRef[]>([]);
  /** The container/member handles' measured geometry (page-relative px). */
  const groupHandles = ref<ReturnType<typeof noImageHandleGeometry>>(noImageHandleGeometry());
  /** The anchor marker for a selected floating container (SPEC-051's
   * `shape-anchor-marker`, attached to the container exactly as to a single
   * shape — SPEC-057 AC-14): page-relative px, or null when the container is
   * inline, a member is selected, or there is no anchor paragraph to mark. */
  const groupAnchorMarkerPx = ref<{ left: number; top: number } | null>(null);

  function beforeSelection(): ModelSelection {
    const root = contentRoot.value as HTMLElement | null;
    return currentSelection(root) ?? caretSelection(getCaret() ?? documentStart());
  }

  function containerAt(section: number, block: number): DrawingGroup | null {
    const b = doc.value?.sections[section]?.blocks[block];
    return isDrawingGroup(b) ? b : null;
  }

  /** The currently selected container block, if a group is selected at any
   * depth. */
  function selectedContainer(): DrawingGroup | null {
    const sel = groupSelection.value;
    return sel ? containerAt(sel.section, sel.block) : null;
  }

  function selectionIsContainer(): boolean {
    const sel = groupSelection.value;
    return !!sel && (!sel.path || sel.path.length === 0);
  }

  function clearGroupSelection() {
    groupSelection.value = null;
    multiSelection.value = [];
    groupHandles.value = null;
    resetMemberTextEditingState();
  }

  // --- selection (click-to-enter, one level per click) ----------------------

  /** Resolves a click on a container's proxy tree (SPEC-057 "Selection across
   * drawing kinds"): the first click selects the outermost container, each
   * further click on the same drawing descends one level along the clicked
   * member path. A sealed container never enters. */
  function handleGroupSelect(section = 0, block = 0, payload = { path: [] as number[], shiftKey: false }) {
    const fullPath = payload?.path ?? [];
    const shiftKey = payload?.shiftKey ?? false;
    const container = containerAt(section, block);
    if (!container) {
      return;
    }
    // Selecting any proxy (never a click inside the live text frame, whose
    // proxy goes pointer-events:none) leaves member text editing.
    resetMemberTextEditingState();
    if (shiftKey) {
      // Shift-click widens the selection to a cross-kind multi-selection
      // (SPEC-051 gesture, widened by SPEC-057 AC-7) — it must NOT wipe the
      // sibling shape/image/chart selections, which join this container as
      // Group candidates. Seed the unified list with the container that was
      // singly selected so a group can be the first item combined.
      if (multiSelection.value.length === 0 && groupSelection.value && (groupSelection.value.path?.length ?? 0) === 0) {
        multiSelection.value = [{ section: groupSelection.value.section, block: groupSelection.value.block }];
      }
      toggleMultiSelection({ section, block });
      return;
    }
    // A plain click is exclusive: it clears the sibling drawing selections.
    clearOtherSelections();
    multiSelection.value = [];
    const sel = groupSelection.value;
    const onThis = !!sel && sel.section === section && sel.block === block;
    const curDepth = onThis ? sel.path?.length ?? 0 : -1;
    let nextDepth = Math.min(curDepth + 1, fullPath.length);
    // A sealed container does not enter: a second click keeps it selected.
    if (isSealedGroup(container)) {
      nextDepth = 0;
    }
    groupSelection.value = { section, block, path: fullPath.slice(0, nextDepth) };
    void nextTick().then(remeasureGroupHandles);
  }

  /** Double-clicking a member (SPEC-057 AC-25): descends selection to it and,
   * when it is a shape member (a grouped text box), enters its text through
   * SPEC-051's own shape-text lifecycle — so text inside a grouped text box is
   * editable in place with every feature that lifecycle gives (typing, lists,
   * tables, IME, Escape-to-leave). A non-shape member (image/chart/nested
   * container) or a sealed container just settles the selection. */
  function handleGroupEditText(section = 0, block = 0, payload = { path: [] as number[] }) {
    const path = payload?.path ?? [];
    const container = containerAt(section, block);
    if (!container || isSealedGroup(container)) {
      return;
    }
    clearOtherSelections();
    multiSelection.value = [];
    groupSelection.value = { section, block, path };
    const model = doc.value;
    const location: DrawingLocation = { section, block, path };
    if (model && path.length > 0 && memberShapeAt(model, location)) {
      enterMemberTextEditing(location);
    } else {
      void nextTick().then(remeasureGroupHandles);
    }
  }

  // --- member text-frame editing (SPEC-057 AC-25) ---------------------------
  // A grouped shape member's text frame is a second live-editable surface, the
  // same one-live-surface pattern useShapes.ts follows for a top-level shape:
  // its own caret/anchor/pending-format/composition state, addressed through
  // withMemberTextModel's synthetic single-section coordinates (section 0) and
  // folded back by applyMemberTextEdit (commitMemberTextModel). The member the
  // path addresses is always the current groupSelection while editing.

  /** True while a shape member's text frame is being edited in place. */
  const editingMemberText = ref(false);
  let memberTextCaret: CaretPosition | null = noCaret();
  let memberTextSelectionAnchor: CaretPosition | null = noCaret();
  let memberTextPendingFormat: ReturnType<typeof noPendingFormat> = null;
  let memberTextComposing = false;
  let memberTextCompositionRange: ModelSelection | null = noSelection();

  /** The one DOM element currently rendering as a member's editable text frame
   * — DocGroupMemberVisual sets data-member-text-editable only on the member
   * whose path is being edited, so at most one exists at a time. */
  function memberTextRootEl(): HTMLElement | null {
    const found = contentRoot.value?.querySelector("[data-member-text-editable]");
    return found instanceof HTMLElement ? found : null;
  }

  function currentMemberTextSelection(): ModelSelection | null {
    const root = memberTextRootEl();
    if (root) {
      const range = modelRangeFromSelection(root, document.getSelection());
      if (range) {
        return { anchor: range.start, focus: range.end };
      }
    }
    return memberTextCaret ? { anchor: memberTextSelectionAnchor ?? memberTextCaret, focus: memberTextCaret } : null;
  }

  function resetMemberTextEditingState(): void {
    editingMemberText.value = false;
    memberTextCaret = noCaret();
    memberTextSelectionAnchor = noCaret();
    memberTextPendingFormat = noPendingFormat();
  }

  /** Leaves member text editing (Escape), keeping the member selected. */
  function exitMemberTextEditing(): void {
    if (!editingMemberText.value) {
      return;
    }
    resetMemberTextEditingState();
    void nextTick().then(() => { (contentRoot.value as HTMLElement | null)?.focus(); remeasureGroupHandles(); });
  }

  /** Puts the caret inside a shape member's text frame (AC-25), populating an
   * empty frame first — the member equivalent of handleShapeEditText. */
  function enterMemberTextEditing(location: DrawingLocation) {
    const model = doc.value;
    if (!model || !editable.value || !memberShapeAt(model, location)) {
      return;
    }
    if (!memberShapeAt(model, location)?.text) {
      const before = beforeSelection();
      applyEdit(ensureMemberTextFrame(model, before, location), before, "add shape text");
    }
    editingMemberText.value = true;
    const current = doc.value;
    const text = current ? memberShapeAt(current, location)?.text : undefined;
    const lastBlock = text ? Math.max(0, text.blocks.length - 1) : 0;
    const lastLength = text ? blockLength(text.blocks[lastBlock]) : 0;
    memberTextCaret = { section: 0, block: lastBlock, offset: lastLength };
    memberTextSelectionAnchor = noCaret();
    memberTextPendingFormat = noPendingFormat();
    void nextTick().then(() => {
      remeasureGroupHandles();
      const root = memberTextRootEl();
      if (root) {
        root.focus();
        setDomCaret(root, memberTextCaret ?? { section: 0, block: 0, offset: 0 });
      }
    });
  }

  /** Folds a `withMemberTextModel` command result back into the real document
   * and commits it — applyShapeTextEdit's sibling (see useShapes.ts). Keeps the
   * caret inside the member frame, so editing continues uninterrupted. */
  function applyMemberTextEdit(resultModel: DocumentModel, resultSelection: ModelSelection, label = ""): void {
    const sel = groupSelection.value;
    const before = doc.value;
    if (!sel || !before || !resultSelection) {
      return;
    }
    const location: DrawingLocation = { section: sel.section, block: sel.block, path: sel.path ?? [] };
    const nextModel = commitMemberTextModel(before, location, resultModel);
    const bodyFiller = caretSelection(getCaret() ?? documentStart());
    recordHistoryEdit(documentId.value, {
      before: { model: before, selection: bodyFiller },
      after: { model: nextModel, selection: bodyFiller },
      label,
      groupable: false,
    });
    doc.value = replaceDocumentModel(documentId.value, nextModel);
    memberTextCaret = resultSelection.focus;
    memberTextSelectionAnchor = isCollapsed(resultSelection) ? noCaret() : resultSelection.anchor;
    markDirty();
    refreshUndoRedoState();
    void nextTick().then(() => {
      const root = memberTextRootEl();
      if (!root || !memberTextCaret) {
        return;
      }
      if (memberTextSelectionAnchor) {
        setDomSelection(root, memberTextSelectionAnchor, memberTextCaret);
      } else {
        setDomCaret(root, memberTextCaret);
      }
    });
  }

  function handleMemberTextBeforeInput(event: InputEvent) {
    const type = event.inputType;
    if (memberTextComposing || event.isComposing || type === "insertCompositionText") {
      return;
    }
    event.preventDefault();
    const sel = groupSelection.value;
    const model = doc.value;
    if (!sel || !model || !editable.value) {
      return;
    }
    const location: DrawingLocation = { section: sel.section, block: sel.block, path: sel.path ?? [] };
    const selection = currentMemberTextSelection();
    if (!selection) {
      return;
    }
    const synthetic = withMemberTextModel(model, location);
    let result: { model: DocumentModel; selection: ModelSelection } | null = null;
    let label = "";
    if (type === "insertText" || type === "insertFromPaste") {
      const text = event.data ?? (event.dataTransfer ? event.dataTransfer.getData("text/plain") : "");
      if (text) {
        result = insertText(synthetic, selection, text, memberTextPendingFormat);
        label = "typing";
      }
    } else if (type === "insertParagraph") {
      result = splitBlock(synthetic, selection);
      label = "new paragraph";
    } else if (type === "insertLineBreak") {
      result = insertLineBreak(synthetic, selection);
      label = "line break";
    } else if (type === "deleteContentBackward") {
      result = deleteBackward(synthetic, selection);
      label = "delete";
    } else if (type === "deleteContentForward") {
      result = deleteForward(synthetic, selection);
      label = "delete";
    } else if (type === "deleteByCut" || type === "deleteContent") {
      result = isCollapsed(selection) ? null : deleteRange(synthetic, selection);
      label = "delete";
    }
    if (result) {
      memberTextPendingFormat = noPendingFormat();
      applyMemberTextEdit(result.model, result.selection, label);
    }
  }

  function handleMemberTextKeyDown(event: KeyboardEvent) {
    if (event.key === "Escape") {
      event.stopPropagation();
      exitMemberTextEditing();
    }
  }

  function handleMemberTextCompositionStart() {
    if (!memberTextRootEl() || !editable.value) {
      return;
    }
    memberTextComposing = true;
    memberTextCompositionRange = currentMemberTextSelection();
  }

  function handleMemberTextCompositionEnd(event: CompositionEvent) {
    memberTextComposing = false;
    const range = memberTextCompositionRange;
    memberTextCompositionRange = null;
    const sel = groupSelection.value;
    const model = doc.value;
    if (!sel || !model || !editable.value || !range) {
      return;
    }
    const data = event.data ?? "";
    if (!data) {
      return;
    }
    const location: DrawingLocation = { section: sel.section, block: sel.block, path: sel.path ?? [] };
    const synthetic = withMemberTextModel(model, location);
    const result = insertText(synthetic, range, data, null);
    applyMemberTextEdit(result.model, result.selection, "typing");
  }

  /** Escape steps back out one level of group selection (SPEC-057 AC-24). */
  function escapeGroupSelection(): boolean {
    const sel = groupSelection.value;
    if (!sel) {
      return false;
    }
    // While editing a member's text, Escape leaves the text (member stays
    // selected) before it steps the selection out a level.
    if (editingMemberText.value) {
      exitMemberTextEditing();
      return true;
    }
    if (sel.path && sel.path.length > 0) {
      groupSelection.value = { section: sel.section, block: sel.block, path: sel.path.slice(0, -1) };
      void nextTick().then(() => { focusSelectedProxy(); remeasureGroupHandles(); });
      return true;
    }
    clearGroupSelection();
    return true;
  }

  function toggleMultiSelection(ref: DrawingRef) {
    const current = multiSelection.value;
    const index = current.findIndex((r) => r.section === ref.section && r.block === ref.block);
    multiSelection.value = index >= 0 ? current.filter((_, i) => i !== index) : [...current, ref];
    groupSelection.value = null;
  }

  /** Adds a top-level drawing to the unified multi-selection (SPEC-057) — the
   * shift-click entry point DocumentView routes every drawing kind through. */
  function addToMultiSelection(section = 0, block = 0) {
    toggleMultiSelection({ section, block });
  }

  /** The drawings Group would combine: SPEC-051's shape Shift-click list
   * (kind-agnostic in storage) unioned with any containers this composable's
   * own Shift-click added, de-duplicated by address. */
  function groupCandidates(): DrawingRef[] {
    const seen = new Set<string>();
    const out: DrawingRef[] = [];
    // A singly-selected top-level container (path empty) is itself a Group
    // candidate (SPEC-057 AC-7) — grouping a selection that includes a group
    // nests that group as one member.
    const single = groupSelection.value;
    const singleContainer: DrawingRef[] =
      single && (single.path?.length ?? 0) === 0 ? [{ section: single.section, block: single.block }] : [];
    for (const r of [...getGroupCandidates(), ...multiSelection.value, ...singleContainer]) {
      const key = `${r.section}:${r.block}`;
      if (!seen.has(key)) {
        seen.add(key);
        out.push({ section: r.section, block: r.block });
      }
    }
    return out;
  }

  function canGroup(): boolean {
    const candidates = groupCandidates();
    if (candidates.length < 2) {
      return false;
    }
    return new Set(candidates.map((r) => r.section)).size === 1;
  }

  // --- DOM measurement (the layout the lib module is kept free of) ----------

  /** The laid-out page rectangle of a top-level drawing, in cm relative to the
   * page — the same geometry the painter used, unioned to build a container. */
  function drawingRectCm(section: number, block: number): RectCm | null {
    const pageEl = pageRef.value;
    const root = contentRoot.value;
    if (!pageEl || !root) {
      return null;
    }
    const anchor = root.querySelector(
      `.doc-shape-anchor[data-section="${section}"][data-block="${block}"], .doc-image-anchor[data-section="${section}"][data-block="${block}"], .doc-chart-anchor[data-section="${section}"][data-block="${block}"], .doc-group-anchor[data-section="${section}"][data-block="${block}"]`,
    );
    const inner = anchor?.querySelector(".doc-shape, .doc-image, .doc-chart, .doc-group") ?? anchor;
    if (!inner) {
      return null;
    }
    const rect = inner.getBoundingClientRect();
    const pageRect = pageEl.getBoundingClientRect();
    return {
      xCm: (rect.left - pageRect.left) * CM_PER_PX,
      yCm: (rect.top - pageRect.top) * CM_PER_PX,
      widthCm: rect.width * CM_PER_PX,
      heightCm: rect.height * CM_PER_PX,
    };
  }

  /** The laid-out page position (cm, page-relative) of a top-level drawing's
   * anchor span — the reference frame `groupDisplayStyle` positions the
   * container within (`.doc-*-anchor`, made `position: relative` when floating).
   * The new container inherits the first member's block slot, so its own anchor
   * lands here; the container's offset is measured against it so Group moves no
   * member on screen (SPEC-057 AC-4). */
  function drawingAnchorCm(section: number, block: number): { xCm: number; yCm: number } | null {
    const pageEl = pageRef.value;
    const root = contentRoot.value;
    if (!pageEl || !root) {
      return null;
    }
    const anchor = root.querySelector(
      `.doc-shape-anchor[data-section="${section}"][data-block="${block}"], .doc-image-anchor[data-section="${section}"][data-block="${block}"], .doc-chart-anchor[data-section="${section}"][data-block="${block}"], .doc-group-anchor[data-section="${section}"][data-block="${block}"]`,
    );
    if (!anchor) {
      return null;
    }
    const rect = anchor.getBoundingClientRect();
    const pageRect = pageEl.getBoundingClientRect();
    return { xCm: (rect.left - pageRect.left) * CM_PER_PX, yCm: (rect.top - pageRect.top) * CM_PER_PX };
  }

  /** Moves DOM focus to the currently selected container/member proxy so
   * keyboard traversal and screen readers follow the selection (SPEC-057
   * "Rendering and test surface"). The proxy testids are built from the
   * container's own `drawing-<index>` base, which we read off the rendered
   * container proxy. */
  /** The DOM proxy element of the currently selected container/member — the
   * hit/measure surface every handle and focus attaches to. Its testid is the
   * container's own `drawing-<index>` base (read off the rendered container
   * proxy) plus one `-member-<j>` per path step. */
  function selectedProxyEl(): HTMLElement | null {
    const sel = groupSelection.value;
    const root = contentRoot.value;
    if (!sel || !root) {
      return null;
    }
    const anchor = root.querySelector(`.doc-group-anchor[data-section="${sel.section}"][data-block="${sel.block}"]`);
    const containerProxy = anchor?.querySelector(".doc-group-proxy");
    const base = containerProxy?.getAttribute("data-testid");
    if (!base) {
      return null;
    }
    const path = sel.path ?? [];
    const testid = path.length === 0 ? base : base + path.map((j) => `-member-${j}`).join("");
    const el = root.querySelector(`[data-testid="${testid}"]`);
    return el instanceof HTMLElement ? el : null;
  }

  function focusSelectedProxy() {
    selectedProxyEl()?.focus();
  }

  /** The container `.members.length` at the selected path's parent, for Tab
   * traversal, plus whether the selected drawing is itself an enterable
   * container (for Enter). */
  function selectedNode() {
    const container = selectedContainer();
    const sel = groupSelection.value;
    if (!container || !sel) {
      return { node: null, siblingCount: 0, parent: null };
    }
    const path = sel.path ?? [];
    const node = memberAtPath(container, path);
    const parentPath = path.slice(0, -1);
    const parent = memberAtPath(container, parentPath);
    const siblingCount = parent && isDrawingGroup(parent) ? parent.members.length : 0;
    return { node, siblingCount, parent };
  }

  /** Measures the selected container/member's handle geometry (SPEC-057) — the
   * selected proxy is the box every handle attaches to, so members get handles
   * inside their container exactly as the container gets them on the page. */
  function remeasureGroupHandles() {
    const pageEl = pageRef.value;
    const el = selectedProxyEl();
    groupHandles.value = pageEl && el ? measureImageGeometry(pageEl, el) : null;
    const root = contentRoot.value as HTMLElement | null;
    const sel = groupSelection.value;
    groupAnchorMarkerPx.value =
      pageEl && root && sel && selectionIsContainer() ? measureGroupAnchorMarker(sel.section, sel.block, pageEl, root) : null;
  }

  /** The paragraph a container at (section, block) is anchored to — the block
   * immediately before it in the block list, else the one after, else null —
   * mirroring useShapes' anchorParagraphBlockIndex (SPEC-057 AC-14). */
  function containerAnchorParagraphBlockIndex(section: number, block: number): number | null {
    const blocks = doc.value?.sections[section]?.blocks;
    if (!blocks) {
      return null;
    }
    if (blocks[block - 1]?.type === "paragraph") {
      return block - 1;
    }
    if (blocks[block + 1]?.type === "paragraph") {
      return block + 1;
    }
    return null;
  }

  /** The DOM position of the paragraph a selected container is anchored to
   * (SPEC-057 AC-14); null when there is none, or it isn't measurable. */
  function measureGroupAnchorMarker(section: number, block: number, pageEl: Element, root: Element): { left: number; top: number } | null {
    const anchorBlockIndex = containerAnchorParagraphBlockIndex(section, block);
    if (anchorBlockIndex === null) {
      return null;
    }
    const el = root.querySelector(`[data-section="${section}"][data-block="${anchorBlockIndex}"]`);
    if (!el) {
      return null;
    }
    const pageRect = pageEl.getBoundingClientRect();
    const rect = el.getBoundingClientRect();
    return { left: rect.left - pageRect.left, top: rect.top - pageRect.top };
  }

  /** Every paragraph in the container's section, as re-anchoring candidates
   * (page-relative cm) — the container is a top-level block, so every paragraph
   * is a valid candidate, mirroring useShapes' candidateParagraphsFor. */
  function containerCandidateParagraphs(section: number): ParagraphSpan[] {
    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    const pageEl = pageRef.value;
    if (!model || !root || !pageEl) {
      return [];
    }
    const pageRect = pageEl.getBoundingClientRect();
    const blocks = model.sections[section]?.blocks ?? [];
    const spans: ParagraphSpan[] = [];
    blocks.forEach((block, index) => {
      if (block.type !== "paragraph") {
        return;
      }
      const el = root.querySelector(`[data-section="${section}"][data-block="${index}"]`);
      if (!el) {
        return;
      }
      const rect = el.getBoundingClientRect();
      spans.push({ section, block: index, topCm: (rect.top - pageRect.top) * CM_PER_PX, bottomCm: (rect.bottom - pageRect.top) * CM_PER_PX });
    });
    return spans;
  }

  // --- container / member resize + rotation drag ----------------------------
  // Both commit once, on release (SPEC-057 "Undo"): a mid-drag preview turns
  // the content layer / resizes the box via CSS, the model changes only at
  // mouseup. Resize/rotate are the same handles a single shape uses, so a
  // member is resized/rotated inside its container exactly as on the page.

  /** The selected drawing (container when the path is empty, else the member
   * the path addresses). */
  function selectedDrawing(): DrawingMember | null {
    const container = selectedContainer();
    const sel = groupSelection.value;
    if (!container || !sel) {
      return null;
    }
    return memberAtPath(container, sel.path ?? []);
  }

  interface ResizeDrag {
    handle: string;
    startClientX: number;
    startClientY: number;
    startWidthCm: number;
    startHeightCm: number;
    rotationDeg: number;
    widthCm: number;
    heightCm: number;
  }
  let resizeDrag: ResizeDrag | null = null;
  let resizeMoved = false;

  interface RotateDrag {
    centerClientX: number;
    centerClientY: number;
    startAngle: number;
    startRotation: number;
    contentEl: HTMLElement | null;
  }
  let rotateDrag: RotateDrag | null = null;
  let rotateMoved = false;

  /** The rotated content layer of the selected drawing (for live rotation
   * preview): the container's own `.doc-group-content`, or a member wrapper. */
  function selectedContentEl(): HTMLElement | null {
    const proxy = selectedProxyEl();
    if (!proxy) {
      return null;
    }
    // Container proxy → its sibling content layer; member proxy → its wrapper.
    const content = proxy.parentElement?.querySelector(":scope > .doc-group-content");
    return content instanceof HTMLElement ? content : (proxy.parentElement instanceof HTMLElement ? proxy.parentElement : null);
  }

  function startGroupResize(event: MouseEvent, handle = "se") {
    const drawing = selectedDrawing();
    if (!drawing) {
      return;
    }
    event.preventDefault();
    resizeMoved = false;
    resizeDrag = {
      handle,
      startClientX: event.clientX,
      startClientY: event.clientY,
      startWidthCm: drawing.widthCm,
      startHeightCm: drawing.heightCm,
      rotationDeg: drawing.rotationDeg ?? 0,
      widthCm: drawing.widthCm,
      heightCm: drawing.heightCm,
    };
    window.addEventListener("mousemove", handleGroupResizeMove);
    window.addEventListener("mouseup", finishGroupResize);
  }

  function handleGroupResizeMove(event: MouseEvent) {
    const drag = resizeDrag;
    if (!drag) {
      return;
    }
    resizeMoved = true;
    const local = unrotateDelta(
      (event.clientX - drag.startClientX) * CM_PER_PX,
      (event.clientY - drag.startClientY) * CM_PER_PX,
      drag.rotationDeg,
    );
    const size = resizedImageSize({ widthCm: drag.startWidthCm, heightCm: drag.startHeightCm }, asImageHandle(drag.handle), local.dx, local.dy, event.shiftKey);
    drag.widthCm = size.widthCm;
    drag.heightCm = size.heightCm;
  }

  function finishGroupResize() {
    window.removeEventListener("mousemove", handleGroupResizeMove);
    window.removeEventListener("mouseup", finishGroupResize);
    const drag = resizeDrag;
    resizeDrag = null;
    const sel = groupSelection.value;
    const model = doc.value;
    if (!drag || !resizeMoved || !sel || !model) {
      return;
    }
    const before = beforeSelection();
    if (selectionIsContainer()) {
      applyEdit(resizeGroupContainer(model, before, sel.section, sel.block, drag.widthCm, drag.heightCm), before, "resize group");
    } else {
      applyEdit(resizeMember(model, before, sel, drag.widthCm, drag.heightCm), before, "edit group member");
    }
    void nextTick().then(remeasureGroupHandles);
  }

  function startGroupRotate(event: MouseEvent) {
    const drawing = selectedDrawing();
    const proxy = selectedProxyEl();
    if (!drawing || !proxy) {
      return;
    }
    event.preventDefault();
    const rect = proxy.getBoundingClientRect();
    const cx = rect.left + rect.width / 2;
    const cy = rect.top + rect.height / 2;
    rotateMoved = false;
    rotateDrag = {
      centerClientX: cx,
      centerClientY: cy,
      startAngle: (Math.atan2(event.clientY - cy, event.clientX - cx) * 180) / Math.PI,
      startRotation: drawing.rotationDeg ?? 0,
      contentEl: selectedContentEl(),
    };
    window.addEventListener("mousemove", handleGroupRotateMove);
    window.addEventListener("mouseup", finishGroupRotate);
  }

  let rotatePreviewDeg = 0;

  function handleGroupRotateMove(event: MouseEvent) {
    const drag = rotateDrag;
    if (!drag) {
      return;
    }
    rotateMoved = true;
    const angle = (Math.atan2(event.clientY - drag.centerClientY, event.clientX - drag.centerClientX) * 180) / Math.PI;
    rotatePreviewDeg = Math.round(drag.startRotation + (angle - drag.startAngle));
    // Live preview: turn the content layer without touching the model.
    if (drag.contentEl) {
      const prop = selectionIsContainer() ? "--group-rotation" : "--member-rotation";
      drag.contentEl.style.setProperty(prop, `${rotatePreviewDeg}deg`);
    }
  }

  function finishGroupRotate() {
    window.removeEventListener("mousemove", handleGroupRotateMove);
    window.removeEventListener("mouseup", finishGroupRotate);
    const drag = rotateDrag;
    rotateDrag = null;
    const sel = groupSelection.value;
    const model = doc.value;
    if (!drag || !rotateMoved || !sel || !model) {
      return;
    }
    const before = beforeSelection();
    if (selectionIsContainer()) {
      applyEdit(rotateGroupContainer(model, before, sel.section, sel.block, rotatePreviewDeg), before, "rotate group");
    } else {
      applyEdit(rotateMember(model, before, sel, rotatePreviewDeg), before, "edit group member");
    }
    void nextTick().then(remeasureGroupHandles);
  }

  /** Drags the selected container/member's body to move it (SPEC-057
   * AC-10/AC-26); commits once on release. Armed by useCanvasGestures on a
   * `.doc-group-anchor` mousedown while a group is selected. Not while the
   * member's text frame is live (editingMemberText) — the frame stops
   * beforeinput/keydown but not mousedown, and a drag there is text
   * selection, not a move (startShapeMove's editingShapeText guard). */
  let moveDrag: { startClientX: number; startClientY: number } | null = null;
  let moveMoved = false;

  function startGroupMove(event: MouseEvent) {
    if (!groupSelection.value || editingMemberText.value) {
      return;
    }
    moveMoved = false;
    moveDrag = { startClientX: event.clientX, startClientY: event.clientY };
    window.addEventListener("mousemove", handleGroupMoveMove);
    window.addEventListener("mouseup", finishGroupMove);
  }

  function handleGroupMoveMove(event: MouseEvent) {
    if (!moveDrag) {
      return;
    }
    if (Math.abs(event.clientX - moveDrag.startClientX) > 2 || Math.abs(event.clientY - moveDrag.startClientY) > 2) {
      moveMoved = true;
    }
  }

  function finishGroupMove(event: MouseEvent) {
    window.removeEventListener("mousemove", handleGroupMoveMove);
    window.removeEventListener("mouseup", finishGroupMove);
    const drag = moveDrag;
    moveDrag = null;
    const sel = groupSelection.value;
    const model = doc.value;
    if (!drag || !moveMoved || !sel || !model) {
      return;
    }
    const dxCm = (event.clientX - drag.startClientX) * CM_PER_PX;
    const dyCm = (event.clientY - drag.startClientY) * CM_PER_PX;
    const before = beforeSelection();
    if (selectionIsContainer()) {
      const moved = moveGroupContainer(model, before, sel.section, sel.block, dxCm, dyCm);
      // Re-anchor across a paragraph boundary on top of the move (SPEC-057
      // AC-14) — the container is one drawing, so this is SPEC-051's shape
      // re-anchoring applied to the container's own box.
      const candidates = containerCandidateParagraphs(sel.section);
      const anchorBlockIndex = containerAnchorParagraphBlockIndex(sel.section, sel.block);
      const anchorSpan = anchorBlockIndex === null ? undefined : candidates.find((span) => span.block === anchorBlockIndex);
      const rect = drawingRectCm(sel.section, sel.block);
      let result = moved;
      let nextBlock = sel.block;
      if (anchorSpan && rect) {
        result = reanchorShape(moved.model, before, { section: sel.section, block: sel.block }, candidates, rect.yCm + dyCm + rect.heightCm / 2, anchorSpan.topCm);
        nextBlock = result.selection.focus.block - 1;
      }
      applyEdit(result, before, "move group");
      groupSelection.value = { section: sel.section, block: nextBlock, path: [] };
    } else {
      applyEdit(moveMember(model, before, sel, dxCm, dyCm), before, "edit group member");
    }
    void nextTick().then(remeasureGroupHandles);
  }

  // --- Group / Ungroup ------------------------------------------------------

  /** Groups the unified multi-selection into one container (SPEC-057 AC-3): the
   * gesture layer supplies each member's laid-out rectangle so the pure op
   * stays layout-free. Leaves the new container selected. */
  function groupSelectedDrawings() {
    const model = doc.value;
    if (!model || !canGroup()) {
      return;
    }
    const entries: GroupEntry[] = [];
    for (const ref of groupCandidates()) {
      const rect = drawingRectCm(ref.section, ref.block);
      if (!rect) {
        return;
      }
      entries.push({ section: ref.section, block: ref.block, rect });
    }
    const section = entries[0].section;
    const firstBlock = Math.min(...entries.map((e) => e.block));
    // The container inherits the first member's block slot; measure that anchor
    // so its offset comes out anchor-relative and no member jumps (AC-4).
    const anchor = drawingAnchorCm(section, firstBlock) ?? { xCm: 0, yCm: 0 };
    const before = beforeSelection();
    const result = groupDrawings(model, before, entries, anchor);
    applyEdit(result, before, "group drawings");
    multiSelection.value = [];
    clearOtherSelections();
    groupSelection.value = { section, block: firstBlock, path: [] };
    void nextTick().then(remeasureGroupHandles);
  }

  /** Ungroups the selected container one level (SPEC-057 AC-5): the emerged
   * members become the multi-selection so Ctrl/Cmd+G re-groups immediately. */
  function ungroupSelectedContainer() {
    const model = doc.value;
    const sel = groupSelection.value;
    if (!model || !sel) {
      return;
    }
    const before = beforeSelection();
    const result = ungroupDrawing(model, before, sel.section, sel.block);
    applyEdit(result, before, "ungroup drawings");
    groupSelection.value = null;
    multiSelection.value = (result.emerged ?? []).map((e) => ({ section: e.section, block: e.block }));
    void nextTick().then(remeasureGroupHandles);
  }

  // --- container / member commands -----------------------------------------

  /** Delete: removes the selected member, or the whole container (SPEC-057
   * AC-12/AC-31). A container reduced to one member dissolves. */
  function deleteSelected() {
    const model = doc.value;
    const sel = groupSelection.value;
    if (!model || !sel) {
      return;
    }
    const before = beforeSelection();
    if (selectionIsContainer()) {
      applyEdit(deleteGroupContainer(model, before, sel.section, sel.block), before, "delete group");
      clearGroupSelection();
    } else {
      applyEdit(deleteMember(model, before, sel), before, "delete group member");
      // A container reduced to one member has dissolved; re-select the block.
      groupSelection.value = containerAt(sel.section, sel.block) ? { section: sel.section, block: sel.block, path: [] } : null;
      void nextTick().then(remeasureGroupHandles);
    }
  }

  /** Ctrl/Cmd+D: duplicates the selected member or container (SPEC-057). */
  function duplicateSelected() {
    const model = doc.value;
    const sel = groupSelection.value;
    if (!model || !sel) {
      return;
    }
    const before = beforeSelection();
    if (selectionIsContainer()) {
      applyEdit(duplicateGroupContainer(model, before, sel.section, sel.block), before, "duplicate group");
      groupSelection.value = { section: sel.section, block: sel.block + 1, path: [] };
    } else {
      applyEdit(duplicateMember(model, before, sel), before, "duplicate group member");
      groupSelection.value = { section: sel.section, block: sel.block, path: duplicatedMemberPath(sel) };
    }
    void nextTick().then(remeasureGroupHandles);
  }

  /** Arrow-key nudge of the selected container or member (SPEC-057
   * AC-11/AC-27). */
  function nudgeSelected(direction: NudgeDirection, fine: boolean) {
    const model = doc.value;
    const sel = groupSelection.value;
    if (!model || !sel) {
      return;
    }
    const before = beforeSelection();
    if (selectionIsContainer()) {
      applyEdit(nudgeGroupContainer(model, before, sel.section, sel.block, direction, fine), before, "move group");
    } else {
      applyEdit(nudgeMember(model, before, sel, direction, fine), before, "edit group member");
    }
    void nextTick().then(remeasureGroupHandles);
  }

  /** The four reorder verbs on a selected member (SPEC-057 AC-33..36). */
  function reorderSelectedMember(command: MemberReorder) {
    const model = doc.value;
    const sel = groupSelection.value;
    if (!model || !sel || selectionIsContainer()) {
      return;
    }
    const before = beforeSelection();
    applyEdit(reorderMember(model, before, sel, command), before, "reorder group member");
    void nextTick().then(remeasureGroupHandles);
  }

  /** Rotation from the toolbar's numeric input (SPEC-057 AC-20/AC-29). */
  function setSelectedRotation(deg: number) {
    const model = doc.value;
    const sel = groupSelection.value;
    if (!model || !sel || !Number.isFinite(deg)) {
      return;
    }
    const before = beforeSelection();
    if (selectionIsContainer()) {
      applyEdit(rotateGroupContainer(model, before, sel.section, sel.block, deg), before, "rotate group");
    } else {
      applyEdit(rotateMember(model, before, sel, deg), before, "edit group member");
    }
    void nextTick().then(remeasureGroupHandles);
  }

  /** Alt text from the toolbar (SPEC-057 AC-49) — the container's own, or a
   * selected member's in place. Read-only for a sealed container (the caller
   * hides the input). */
  function setSelectedAltText(altText: string) {
    const model = doc.value;
    const sel = groupSelection.value;
    if (!model || !sel) {
      return;
    }
    const before = beforeSelection();
    if (selectionIsContainer()) {
      applyEdit(setGroupAltText(model, before, sel.section, sel.block, altText), before, "shape alt text");
    } else {
      applyEdit(setMemberAltText(model, before, sel, altText), before, "shape alt text");
    }
  }

  /** Wrap mode from the container toolbar (SPEC-057). */
  function setSelectedWrap(wrap: string) {
    const model = doc.value;
    const sel = groupSelection.value;
    if (!model || !sel || !selectionIsContainer()) {
      return;
    }
    const before = beforeSelection();
    applyEdit(setGroupWrap(model, before, sel.section, sel.block, wrap as DrawingMember["wrap"]), before, "shape wrap");
    void nextTick().then(remeasureGroupHandles);
  }

  /** Z-order from the container toolbar (SPEC-057 AC-16) — the container is one
   * drawing, so it reorders against every sibling drawing through the same
   * `reorderShapeZ` path a single shape uses, reusing SPEC-051's label. */
  function reorderSelectedContainerZ(command: string) {
    const model = doc.value;
    const sel = groupSelection.value;
    if (!model || !sel || !selectionIsContainer()) {
      return;
    }
    const before = beforeSelection();
    applyEdit(reorderShapeZ(model, before, { section: sel.section, block: sel.block }, command as ZOrderCommand), before, "reorder shape");
    void nextTick().then(remeasureGroupHandles);
  }

  /** Ctrl/Cmd+G groups the multi-selection; Ctrl/Cmd+Shift+G ungroups the
   * selected container (SPEC-057). Returns whether it handled the event. */
  function handleGroupShortcut(event: KeyboardEvent): boolean {
    if (!(event.ctrlKey || event.metaKey) || event.key.toLowerCase() !== "g") {
      return false;
    }
    event.preventDefault();
    if (event.shiftKey) {
      if (groupSelection.value && selectionIsContainer()) {
        ungroupSelectedContainer();
      }
    } else if (canGroup()) {
      groupSelectedDrawings();
    }
    return true;
  }

  /** A selected container/member's keyboard interactions (SPEC-057): Delete,
   * Ctrl/Cmd+D duplicate, and arrow-key nudge. Returns whether it handled the
   * event; the caller only calls it while a group is selected. */
  function handleGroupKeyDown(event: KeyboardEvent): boolean {
    const sel = groupSelection.value;
    if (!sel) {
      return false;
    }
    if (event.key === "Enter") {
      // Enter descends into the focused group proxy (SPEC-057 AC-41).
      const { node } = selectedNode();
      if (node && isDrawingGroup(node) && node.members.length > 0) {
        event.preventDefault();
        groupSelection.value = { section: sel.section, block: sel.block, path: [...(sel.path ?? []), 0] };
        void nextTick().then(() => { focusSelectedProxy(); remeasureGroupHandles(); });
        return true;
      }
      return false;
    }
    if (event.key === "Tab" && sel.path && sel.path.length > 0) {
      // Tab traverses siblings within a group proxy (SPEC-057 AC-43), wrapping.
      const { siblingCount } = selectedNode();
      if (siblingCount > 1) {
        event.preventDefault();
        const path = sel.path;
        const i = path[path.length - 1];
        const next = event.shiftKey ? (i - 1 + siblingCount) % siblingCount : (i + 1) % siblingCount;
        groupSelection.value = { section: sel.section, block: sel.block, path: [...path.slice(0, -1), next] };
        void nextTick().then(() => { focusSelectedProxy(); remeasureGroupHandles(); });
        return true;
      }
      return false;
    }
    if (event.key === "Delete" || event.key === "Backspace") {
      event.preventDefault();
      deleteSelected();
      return true;
    }
    if ((event.ctrlKey || event.metaKey) && !event.shiftKey && event.key.toLowerCase() === "d") {
      event.preventDefault();
      duplicateSelected();
      return true;
    }
    const direction =
      event.key === "ArrowUp" ? "up" :
      event.key === "ArrowDown" ? "down" :
      event.key === "ArrowLeft" ? "left" :
      event.key === "ArrowRight" ? "right" : "";
    if (direction) {
      event.preventDefault();
      nudgeSelected(asNudgeDirection(direction), event.ctrlKey || event.metaKey);
      return true;
    }
    return false;
  }

  /** The single keyboard entry point for DocumentView (SPEC-057): Group /
   * Ungroup shortcuts, then a selected container/member's Delete / Ctrl+D /
   * arrow-nudge. Returns whether it handled the event. */
  function handleGroupKeys(event: KeyboardEvent): boolean {
    if (handleGroupShortcut(event)) {
      return true;
    }
    return handleGroupKeyDown(event);
  }

  /** The floating toolbar's position for the selected container/member. */
  function groupToolbarStyle() {
    const handles = groupHandles.value;
    if (!handles) {
      return { left: 0, top: 0, placement: "below" };
    }
    return { left: handles.left, top: Math.max(0, handles.top - 44), placement: "above" };
  }

  /** The selected drawing's rotation (handles turn with it). */
  function selectedRotationDeg(): number {
    return selectedDrawing()?.rotationDeg ?? 0;
  }

  /** Whether the selected drawing offers resize/rotation handles — a sealed
   * container does not (SPEC-057 AC-72/73); its preserved geometry must not be
   * re-coordinated. Members and unsealed containers do. */
  function selectedHasHandles(): boolean {
    if (!groupSelection.value || !groupHandles.value) {
      return false;
    }
    const container = selectedContainer();
    return !(selectionIsContainer() && container !== null && isSealedGroup(container));
  }

  /** The toolbar's full prop bag (SPEC-057), or null when nothing is selected
   * or its handles are not yet measured — so DocumentView renders
   * `<GroupToolbar v-if=state v-bind=state />` in one line. A selected
   * container shows wrap/alt-text/Ungroup (rotation off + alt text read-only
   * when sealed, AC-49/73/74); a selected member shows the four reorder verbs
   * in place of wrap and z-order (AC-33..36). */
  function groupToolbarState() {
    const sel = groupSelection.value;
    if (!sel || !groupHandles.value) {
      return null;
    }
    const pos = groupToolbarStyle();
    const base = { left: pos.left, top: pos.top, placement: pos.placement };
    if (selectionIsContainer()) {
      const container = selectedContainer();
      if (!container) {
        return null;
      }
      return {
        ...base,
        member: false,
        wrap: container.wrap,
        rotationDeg: container.rotationDeg ?? 0,
        altText: container.altText ?? "",
        sealed: isSealedGroup(container),
      };
    }
    const member = selectedDrawing();
    if (!member) {
      return null;
    }
    return {
      ...base,
      member: true,
      wrap: "square",
      rotationDeg: member.rotationDeg ?? 0,
      altText: member.altText ?? "",
      sealed: false,
    };
  }

  /** Selection plus the member being text-edited, for the render layer: the
   * pagination view derives both the selected member path and the editing path
   * (which member's text frame goes contenteditable) from this one object, so
   * the memoized group entry re-renders when either changes. */
  const groupRenderSelection = computed(() =>
    groupSelection.value
      ? { ...groupSelection.value, editingPath: editingMemberText.value ? groupSelection.value.path ?? [] : null }
      : null,
  );

  /** The DocPage listener bag DocumentView binds with a single `v-on` — the
   * container/member selection and double-click-to-edit gestures plus the
   * four member text-frame edit events (AC-25). Keys are the camelCased event
   * names DocPage emits (group-select → groupSelect, …). */
  const groupListeners = {
    groupSelect: handleGroupSelect,
    groupEditText: handleGroupEditText,
    groupTextBeforeinput: handleMemberTextBeforeInput,
    groupTextKeydown: handleMemberTextKeyDown,
    groupTextCompositionstart: handleMemberTextCompositionStart,
    groupTextCompositionend: handleMemberTextCompositionEnd,
  };

  return {
    groupSelection,
    groupRenderSelection,
    groupListeners,
    editingMemberText,
    multiSelection,
    groupHandles,
    clearGroupSelection,
    handleGroupSelect,
    handleGroupEditText,
    escapeGroupSelection,
    addToMultiSelection,
    canGroup,
    groupSelectedDrawings,
    ungroupSelectedContainer,
    deleteSelected,
    duplicateSelected,
    nudgeSelected,
    reorderSelectedMember,
    reorderSelectedContainerZ,
    groupAnchorMarker: () => groupAnchorMarkerPx.value,
    setSelectedRotation,
    setSelectedAltText,
    setSelectedWrap,
    remeasureGroupHandles,
    groupToolbarStyle,
    groupToolbarState,
    selectionIsContainer,
    selectedRotationDeg,
    selectedHasHandles,
    startGroupResize,
    startGroupRotate,
    startGroupMove,
    handleGroupKeys,
  };
}
