// Native drawing-group containers (SPEC-057): selection (click-to-enter,
// second-click-a-member, Escape-out), Group/Ungroup, the container and member
// gestures that reuse the pure operations in src/lib/drawingGroups.ts, the
// container/member toolbar actions, and keyboard proxy traversal. The pure
// model math lives in the lib module; this composable owns only the view
// state and the DOM measurement the layout-free lib deliberately does not.
// Runtime-safe TS (this is a .ts composable, not a .vue script) — see
// useShapes.ts, the sibling this mirrors.
import { nextTick, ref } from "vue";
import {
  caretSelection,
  documentStart,
  type ModelSelection,
} from "../../lib/commands";
import { CM_PER_PX } from "../../lib/core/units";
import { measureImageGeometry, noImageHandleGeometry } from "../../lib/caretGeometry";
import {
  isDrawingGroup,
  isSealedGroup,
  type DrawingGroup,
  type DrawingMember,
} from "../../lib/documentModel";
import {
  deleteGroupContainer,
  deleteMember,
  duplicateGroupContainer,
  duplicateMember,
  duplicatedMemberPath,
  groupDrawings,
  memberAtPath,
  nudgeGroupContainer,
  nudgeMember,
  noDrawingLocation,
  reorderMember,
  rotateGroupContainer,
  rotateMember,
  setGroupAltText,
  setGroupWrap,
  ungroupDrawing,
  type DrawingLocation,
  type GroupEntry,
  type MemberReorder,
  type RectCm,
} from "../../lib/drawingGroups";
import { type EditorContext } from "./context";
import { asNudgeDirection, type NudgeDirection } from "../../lib/core/shapes";

/** What DocumentView passes in beyond the shared context: the ways to clear
 * the sibling per-kind selections (so a group selection is exclusive). */
export interface DrawingGroupsDeps extends EditorContext {
  clearOtherSelections: () => void;
}

/** A top-level drawing address in the unified multi-selection (SPEC-057
 * "Selection across drawing kinds") — shapes, images, charts and containers
 * all join one selection so Group spans kinds. */
interface DrawingRef {
  section: number;
  block: number;
}

export function useDrawingGroups(deps: DrawingGroupsDeps) {
  const { doc, contentRoot, pageRef, currentSelection, applyEdit, getCaret, clearOtherSelections } = deps;

  /** The selected container or member (SPEC-057): `path` empty = the
   * container itself, non-empty = a member by index path. */
  const groupSelection = ref<DrawingLocation | null>(noDrawingLocation());
  /** The unified multi-selection of top-level drawings, for Group. */
  const multiSelection = ref<DrawingRef[]>([]);
  /** The container/member handles' measured geometry (page-relative px). */
  const groupHandles = ref<ReturnType<typeof noImageHandleGeometry>>(noImageHandleGeometry());

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
    clearOtherSelections();
    if (shiftKey) {
      toggleMultiSelection({ section, block });
      return;
    }
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

  /** Double-click composes from clicks (SPEC-057): on a text-bearing member it
   * would enter the member's text; for now it descends selection to the
   * member so the member is at least reachable. */
  function handleGroupEditText(section = 0, block = 0, payload = { path: [] as number[] }) {
    handleGroupSelect(section, block, { path: payload?.path ?? [], shiftKey: false });
  }

  /** Escape steps back out one level of group selection (SPEC-057 AC-24). */
  function escapeGroupSelection(): boolean {
    const sel = groupSelection.value;
    if (!sel) {
      return false;
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

  function canGroup(): boolean {
    if (multiSelection.value.length < 2) {
      return false;
    }
    const sections = new Set(multiSelection.value.map((r) => r.section));
    return sections.size === 1;
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

  /** Moves DOM focus to the currently selected container/member proxy so
   * keyboard traversal and screen readers follow the selection (SPEC-057
   * "Rendering and test surface"). The proxy testids are built from the
   * container's own `drawing-<index>` base, which we read off the rendered
   * container proxy. */
  function focusSelectedProxy() {
    const sel = groupSelection.value;
    const root = contentRoot.value;
    if (!sel || !root) {
      return;
    }
    const anchor = root.querySelector(`.doc-group-anchor[data-section="${sel.section}"][data-block="${sel.block}"]`);
    const containerProxy = anchor?.querySelector(".doc-group-proxy");
    const base = containerProxy?.getAttribute("data-testid");
    if (!base) {
      return;
    }
    const path = sel.path ?? [];
    const testid = path.length === 0 ? base : base + path.map((j) => `-member-${j}`).join("");
    const el = root.querySelector(`[data-testid="${testid}"]`);
    if (el instanceof HTMLElement) {
      el.focus();
    }
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

  function remeasureGroupHandles() {
    const sel = groupSelection.value;
    const pageEl = pageRef.value;
    const root = contentRoot.value;
    if (!sel || !pageEl || !root) {
      groupHandles.value = null;
      return;
    }
    const el = root.querySelector(
      `.doc-group-anchor[data-section="${sel.section}"][data-block="${sel.block}"] .doc-group`,
    );
    groupHandles.value = el ? measureImageGeometry(pageEl, el) : null;
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
    for (const ref of multiSelection.value) {
      const rect = drawingRectCm(ref.section, ref.block);
      if (!rect) {
        return;
      }
      entries.push({ section: ref.section, block: ref.block, rect });
    }
    const before = beforeSelection();
    const result = groupDrawings(model, before, entries);
    applyEdit(result, before, "group drawings");
    const section = entries[0].section;
    const firstBlock = Math.min(...entries.map((e) => e.block));
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

  /** Alt text from the container toolbar (SPEC-057 AC-49). Read-only for a
   * sealed container — the caller hides the input in that case. */
  function setSelectedAltText(altText: string) {
    const model = doc.value;
    const sel = groupSelection.value;
    if (!model || !sel || !selectionIsContainer()) {
      return;
    }
    const before = beforeSelection();
    applyEdit(setGroupAltText(model, before, sel.section, sel.block, altText), before, "shape alt text");
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

  /** The selected container's own model (for the toolbar), or null. */
  function selectedGroupBlock(): DrawingGroup | null {
    return selectionIsContainer() ? selectedContainer() : null;
  }

  /** The container toolbar's full prop bag (SPEC-057 AC-49), or null when no
   * container is selected or its handles are not yet measured — so
   * DocumentView renders `<GroupToolbar v-if=state v-bind=state />` in one
   * line. A sealed container shows rotation off and alt text read-only. */
  function groupToolbarState() {
    const container = selectedGroupBlock();
    if (!container || !groupHandles.value) {
      return null;
    }
    const pos = groupToolbarStyle();
    return {
      left: pos.left,
      top: pos.top,
      placement: pos.placement,
      wrap: container.wrap,
      rotationDeg: container.rotationDeg ?? 0,
      altText: container.altText ?? "",
      sealed: isSealedGroup(container),
    };
  }

  return {
    groupSelection,
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
    setSelectedRotation,
    setSelectedAltText,
    setSelectedWrap,
    remeasureGroupHandles,
    groupToolbarStyle,
    groupToolbarState,
    selectionIsContainer,
    handleGroupKeys,
  };
}
