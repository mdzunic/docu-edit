// Deterministic caret navigation (SPEC-011, SPEC-022 AC-13/14, SPEC-030
// AC-20/21). Character/word/line movement, plain Home/End and mouse selection
// are the browser's native caret behaviour over the rendered canvas; the
// document-level jumps and paging that native behaviour is unreliable for are
// implemented here: page up/down, macOS line Home/End, and the cell-boundary /
// column-boundary arrow escapes that guarantee the caret can never be trapped
// in a table cell or a column box.
import { nextTick } from "vue";
import {
  caretFromDomPoint,
  modelTextLengthOf,
  noElement,
  paragraphElementForCaret,
} from "../../lib/domCaret";
import {
  caretAtClientPoint,
  caretClientRect,
  caretOnBoundaryLine,
} from "../../lib/caretGeometry";
import { arrowEscapeCaret } from "../../lib/tables";
import { isTable } from "../../lib/documentModel";
import { blockLength, compareCaret, paragraphAtCaret } from "../../lib/textEditing";
import { documentEnd, documentStart, isCollapsed } from "../../lib/commands";
import { isMacPlatform } from "../../lib/core/platform";
import type { CaretPosition } from "../../lib/textEditing";
import type { EditorContext } from "./context";

export interface PageNavigationDeps extends EditorContext {
  refreshViewportRange: () => void;
}

/** The block-level children of a rendered column box (skipping zero-height
 * section markers, which are not navigation targets). */
function columnBlockChildren(columnEl = noElement()): Element[] {
  if (!columnEl) {
    return [];
  }
  return Array.from(columnEl.children).filter((child) => child.hasAttribute("data-section"));
}

/** The model caret at the start or end of a rendered block element, read
 * from its data-section/data-block/data-slice-start attributes. */
function caretAtBlockEdge(blockEl = noElement(), edge = "start"): CaretPosition | null {
  if (!blockEl) {
    return null;
  }
  const section = Number(blockEl.getAttribute("data-section"));
  const block = Number(blockEl.getAttribute("data-block"));
  if (!Number.isInteger(section) || !Number.isInteger(block)) {
    return null;
  }
  const sliceStart = Number(blockEl.getAttribute("data-slice-start") ?? "0");
  // modelTextLengthOf, not textContent.length: a note reference mark's
  // rendered number stands for a one-character placeholder (SPEC-031).
  const offset = edge === "start" ? sliceStart : sliceStart + modelTextLengthOf(blockEl);
  return { section, block, offset };
}

/** Where the caret should land when a vertical arrow crosses a column
 * boundary, or null when it should not. */
function columnEscapeTarget(
  root: HTMLElement,
  focus: CaretPosition,
  direction = "down",
): CaretPosition | null {
  if (!root || !focus) {
    return null;
  }
  const paragraphEl = paragraphElementForCaret(root, focus);
  const columnEl = paragraphEl?.closest(".doc-column") ?? null;
  const pageEl = columnEl?.closest(".page") ?? null;
  if (!paragraphEl || !columnEl || !pageEl) {
    return null;
  }
  const children = columnBlockChildren(columnEl);
  const dataCol = Number(columnEl.getAttribute("data-column"));
  if (direction === "down") {
    if (paragraphEl !== children[children.length - 1] || !caretOnBoundaryLine(root, focus, "last")) {
      return null;
    }
    let targetCol: Element | null = pageEl.querySelector(`.doc-column[data-column="${dataCol + 1}"]`);
    if (!targetCol) {
      targetCol = pageEl.nextElementSibling?.querySelector('.doc-column[data-column="0"]') ?? null;
    }
    const first = targetCol ? columnBlockChildren(targetCol)[0] : null;
    return first ? caretAtBlockEdge(first, "start") : null;
  }
  if (paragraphEl !== children[0] || !caretOnBoundaryLine(root, focus, "first")) {
    return null;
  }
  let targetCol: Element | null = null;
  if (dataCol > 0) {
    targetCol = pageEl.querySelector(`.doc-column[data-column="${dataCol - 1}"]`);
  } else {
    const prevCols = pageEl.previousElementSibling?.querySelectorAll(".doc-column");
    targetCol = prevCols && prevCols.length > 0 ? prevCols[prevCols.length - 1] : null;
  }
  const blocks = targetCol ? columnBlockChildren(targetCol) : [];
  const last = blocks[blocks.length - 1] ?? null;
  return last ? caretAtBlockEdge(last, "end") : null;
}

export function usePageNavigation(deps: PageNavigationDeps) {
  const {
    doc,
    contentRoot,
    editable,
    currentSelection,
    applyCaretMove,
    getCaret,
    getScrollContainer,
    refreshViewportRange,
  } = deps;

  /** The caret's moving end (the DOM selection's focus), which Shift+paging
   * keeps extending; falls back to the tracked caret. */
  function domFocusCaret(root = noElement()): CaretPosition | null {
    if (!root) {
      return getCaret();
    }
    const selection = document.getSelection();
    if (selection && selection.rangeCount > 0 && selection.focusNode) {
      const focus = caretFromDomPoint(root, selection.focusNode, selection.focusOffset);
      if (focus) {
        return focus;
      }
    }
    return getCaret();
  }

  /** PageUp/PageDown: scrolls one viewport and moves the caret to the text
   * position one viewport away from where it was, snapping to the document's
   * start/end when there is nothing further to page onto. */
  async function moveCaretByPage(direction = 1, extend = false) {
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    if (!root || !model) {
      return;
    }
    const from = domFocusCaret(root) ?? { section: 0, block: 0, offset: 0 };
    const rect = caretClientRect(root, from);
    const container = getScrollContainer();
    const pageHeight =
      container && container.clientHeight > 0
        ? container.clientHeight
        : window.innerHeight;
    let target: CaretPosition | null = null;
    if (rect) {
      const before = container ? container.scrollTop : 0;
      if (container) {
        container.scrollTop = before + direction * pageHeight;
        // The page scrolled onto may be a virtualized empty box: realize the
        // new viewport before hit-testing for the caret's landing position.
        refreshViewportRange();
        await nextTick();
      }
      const scrolled = container ? container.scrollTop - before : 0;
      target = caretAtClientPoint(
        root,
        rect.left,
        rect.top + direction * pageHeight - scrolled,
      );
      if (target && compareCaret(target, from) === 0) {
        target = null;
      }
    }
    if (!target) {
      target = direction > 0 ? documentEnd(model) : documentStart();
    }
    applyCaretMove(target, extend);
  }

  /** Home/End — the start and end of the caret's *visual* line, so a wrapped
   * paragraph stops at the wrap and not at the paragraph's ends (SPEC-011
   * AC-12). Windows and Linux browsers bind these themselves and the native
   * move is left alone; macOS binds them to scrolling the view instead, which
   * moves no caret at all, so there the editor has to do it. Hit-testing the
   * caret's own line at the paragraph's left/right edge is what makes it
   * visual: the browser resolves the point to the nearest position on that
   * line, wherever the text happens to wrap. Returns true when consumed. */
  function handleLineEdgeKey(event = new KeyboardEvent("keydown")) {
    const root = contentRoot.value as HTMLElement | null;
    if (!root || !isMacPlatform() || (event.key !== "Home" && event.key !== "End")) {
      return false;
    }
    if (event.ctrlKey || event.metaKey || event.altKey) {
      // Cmd+Home/End is the document-start/end command in the registry.
      return false;
    }
    const from = domFocusCaret(root);
    const rect = from ? caretClientRect(root, from) : null;
    const paragraph = from ? paragraphElementForCaret(root, from) : null;
    if (!from || !rect || !paragraph) {
      return false;
    }
    const bounds = paragraph.getBoundingClientRect();
    const toStart = event.key === "Home";
    const target = caretAtClientPoint(
      root,
      toStart ? bounds.left + 1 : bounds.right - 1,
      rect.top + rect.height / 2,
    );
    if (!target) {
      return false;
    }
    event.preventDefault();
    applyCaretMove(target, event.shiftKey);
    return true;
  }

  /** Deterministic cell-boundary arrows (AC-13): a collapsed caret exactly at
   * the boundary its arrow points across — end of the cell for ArrowRight,
   * start for ArrowLeft, first/last visual line of the first/last paragraph
   * for ArrowUp/ArrowDown — moves straight to the neighbouring cell (or out
   * of the table) instead of relying on the browser. Returns true when the
   * event was consumed. */
  function handleTableArrowKey(event = new KeyboardEvent("keydown")) {
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    if (!root || !model || !editable.value) {
      return false;
    }
    const selection = currentSelection(root);
    if (!selection || !isCollapsed(selection) || !selection.focus.cell) {
      return false;
    }
    const focus = selection.focus;
    const cellBlock = paragraphAtCaret(model, focus);
    const cellPath = focus.cell;
    if (!cellBlock || !cellPath) {
      return false;
    }
    const table = model.sections[focus.section].blocks[focus.block];
    if (!isTable(table)) {
      return false;
    }
    const cellBlocks = table.rows[cellPath.row].cells[cellPath.col].blocks;
    const atFirstParagraph = cellPath.block === 0;
    const atLastParagraph = cellPath.block === cellBlocks.length - 1;
    let direction = "";
    if (event.key === "ArrowRight" && atLastParagraph && focus.offset >= blockLength(cellBlock)) {
      direction = "right";
    } else if (event.key === "ArrowLeft" && atFirstParagraph && focus.offset === 0) {
      direction = "left";
    } else if (event.key === "ArrowDown" && atLastParagraph && caretOnBoundaryLine(root, focus, "last")) {
      direction = "down";
    } else if (event.key === "ArrowUp" && atFirstParagraph && caretOnBoundaryLine(root, focus, "first")) {
      direction = "up";
    }
    if (direction === "") {
      return false;
    }
    event.preventDefault();
    const target = arrowEscapeCaret(
      model,
      focus,
      direction === "up" ? "up" : direction === "down" ? "down" : direction === "left" ? "left" : "right",
    );
    if (target) {
      applyCaretMove(target);
    }
    return true;
  }

  /** After the browser processed an arrow keydown with a collapsed caret in a
   * table cell, checks whether the caret actually moved; when it did not, the
   * model-level escape (lib/tables.arrowEscapeCaret) moves it to the
   * neighbouring cell or out of the table — the AC-14 guarantee that native
   * visual navigation cannot strand the caret. */
  function scheduleTableArrowFallback(key = "") {
    const root = contentRoot.value as HTMLElement | null;
    if (!root || !editable.value) {
      return;
    }
    const selection = currentSelection(root);
    if (!selection || !isCollapsed(selection) || !selection.focus.cell) {
      return;
    }
    const before = selection.focus;
    const direction =
      key === "ArrowUp" ? "up" : key === "ArrowDown" ? "down" : key === "ArrowLeft" ? "left" : "right";
    window.setTimeout(() => {
      const model = doc.value;
      const currentRoot = contentRoot.value as HTMLElement | null;
      if (!model || !currentRoot || !editable.value) {
        return;
      }
      const after = currentSelection(currentRoot);
      if (!after || !isCollapsed(after) || compareCaret(after.focus, before) !== 0) {
        return;
      }
      const target = arrowEscapeCaret(model, before, direction);
      if (target) {
        applyCaretMove(target);
      }
    }, 0);
  }

  /** Deterministic cross-column arrows (AC-20/AC-21). Returns true when the
   * event was consumed. */
  function handleColumnArrowKey(event = new KeyboardEvent("keydown")) {
    const root = contentRoot.value as HTMLElement | null;
    if (!root || !editable.value || (event.key !== "ArrowUp" && event.key !== "ArrowDown")) {
      return false;
    }
    const selection = currentSelection(root);
    if (!selection || !isCollapsed(selection) || selection.focus.cell) {
      return false;
    }
    const target = columnEscapeTarget(root, selection.focus, event.key === "ArrowDown" ? "down" : "up");
    if (!target) {
      return false;
    }
    event.preventDefault();
    applyCaretMove(target);
    return true;
  }

  /** After a native vertical arrow with a collapsed caret in a column made no
   * progress, cross the column boundary (mirrors scheduleTableArrowFallback). */
  function scheduleColumnArrowFallback(key = "") {
    const root = contentRoot.value as HTMLElement | null;
    if (!root || !editable.value || (key !== "ArrowUp" && key !== "ArrowDown")) {
      return;
    }
    const selection = currentSelection(root);
    if (!selection || !isCollapsed(selection) || selection.focus.cell) {
      return;
    }
    const paragraphEl = paragraphElementForCaret(root, selection.focus);
    if (!paragraphEl?.closest(".doc-column")) {
      return;
    }
    const before = selection.focus;
    const direction = key === "ArrowDown" ? "down" : "up";
    window.setTimeout(() => {
      const currentRoot = contentRoot.value as HTMLElement | null;
      if (!currentRoot || !editable.value) {
        return;
      }
      const after = currentSelection(currentRoot);
      if (!after || !isCollapsed(after) || compareCaret(after.focus, before) !== 0) {
        return;
      }
      const target = columnEscapeTarget(currentRoot, before, direction);
      if (target && compareCaret(target, before) !== 0) {
        applyCaretMove(target);
      }
    }, 0);
  }

  return {
    moveCaretByPage,
    handleLineEdgeKey,
    handleTableArrowKey,
    scheduleTableArrowFallback,
    handleColumnArrowKey,
    scheduleColumnArrowFallback,
  };
}
