// Hyperlinks and bookmarks (SPEC-025) and table-of-contents navigation (SPEC-033).
// Owning hyperlink/bookmark dialogs and chips, anchor/index jumps, and the TOC
// insert/update/select actions. The chips and dialogs render from the returned
// refs; `handleContentClick` and the core scroll helpers stay in DocumentView
// (handleContentClick also dispatches note-ref clicks).
import { nextTick, ref } from "vue";
import type { Ref } from "vue";
import {
  anchorCaret,
  applyLink,
  autolinkAt,
  deleteBookmark,
  insertBookmark,
  insertLinkText,
  linkAtCaret,
  linkTargetLabel,
  listBookmarks,
  listHeadings,
  noLinkExtent,
  noLinkTarget,
  removeLink,
  sanitizeHref,
  validateBookmarkName,
} from "../../lib/links";
import { extractFragment, fragmentText } from "../../lib/clipboard";
import {
  caretSelection,
  deleteRange,
  documentStart,
  isCollapsed,
  noCommandResult,
  noSelection,
} from "../../lib/commands";
import { paragraphAt, blockLength } from "../../lib/textEditing";
import { defaultTocOptions, insertTableOfContents, updateTocPageNumbersOnly, updateTocEntirely } from "../../lib/toc";
import { caretClientRect } from "../../lib/caretGeometry";
import type { CaretPosition } from "../../lib/textEditing";
import type { ModelSelection } from "../../lib/commands";
import type { DocumentModel } from "../../lib/documentModel";

/** A model selection spanning a link extent — what Edit/Remove act on. */
function extentSelection(extent = noLinkExtent()): ModelSelection | null {
  if (!extent) {
    return null;
  }
  const address = extent.address;
  const anchor = address.cell
    ? { section: address.section, block: address.block, offset: extent.from, cell: address.cell }
    : { section: address.section, block: address.block, offset: extent.from };
  const focus = address.cell
    ? { section: address.section, block: address.block, offset: extent.to, cell: address.cell }
    : { section: address.section, block: address.block, offset: extent.to };
  return { anchor, focus };
}

export interface LinksDeps {
  doc: Ref<DocumentModel | null>;
  contentRoot: Ref<Element | null>;
  editable: Readonly<Ref<boolean>>;
  currentSelection: (root?: HTMLElement | null) => ModelSelection | null;
  paragraphSelection: (root?: HTMLElement | null) => ModelSelection;
  applyEdit: (
    result: { model: DocumentModel; selection: ModelSelection } | null,
    beforeSelection: ModelSelection | null,
    label: string,
    groupable?: boolean,
  ) => void;
  applyCaretMove: (target?: CaretPosition, extend?: boolean) => void;
  /** Core caret-scroll helpers that remain in DocumentView. */
  scrollModelCaretIntoView: (position?: CaretPosition, always?: boolean) => Promise<void>;
  restoreCaretAfterRender: () => Promise<void>;
  /** A TOC is body content — unavailable while a header/footer band is active. */
  activeHeaderFooter: Ref<any>;
  pageOfCaret: (caret: CaretPosition) => number;
  pageNumbers: Ref<number[]>;
  clearCellSelection: () => void;
  clearImageSelection: () => void;
  getCaret: () => CaretPosition | null;
  getSelectionAnchor: () => CaretPosition | null;
}

export function useLinks(deps: LinksDeps) {
  const {
    doc,
    contentRoot,
    editable,
    currentSelection,
    paragraphSelection,
    applyEdit,
    applyCaretMove,
    scrollModelCaretIntoView,
    restoreCaretAfterRender,
    activeHeaderFooter,
    pageOfCaret,
    pageNumbers,
    clearCellSelection,
    clearImageSelection,
    getCaret,
    getSelectionAnchor,
  } = deps;

  const linkDialogOpen = ref(false);
  const linkDialogInitial = ref({ text: "", url: "", anchor: "" });
  let linkDialogSelection = noSelection();
  const linkChip = ref({ visible: false, left: 0, top: 0, label: "", external: false });
  let linkChipExtent = noLinkExtent();
  const bookmarkDialogOpen = ref(false);
  let bookmarkDialogSelection = noSelection();
  const bookmarkManagerOpen = ref(false);

  const tocSelection = ref<{ section: number; blockIndex: number } | null>(null);
  const tocOptionsDialogOpen = ref(false);

  function clearTocSelection() {
    tocSelection.value = null;
  }

  // --- dialogs' target lists ----------------------------------------------

  function documentHeadings() {
    const model = doc.value;
    return model
      ? listHeadings(model).map((heading) => ({
          anchor: heading.anchor,
          text: heading.text,
          level: heading.level,
        }))
      : [];
  }

  function documentBookmarkTargets() {
    const model = doc.value;
    return model
      ? listBookmarks(model).map((bookmark) => ({ anchor: bookmark.name, text: bookmark.name }))
      : [];
  }

  function documentBookmarkNames() {
    const model = doc.value;
    return model ? listBookmarks(model).map((bookmark) => bookmark.name) : [];
  }

  // --- hyperlinks (SPEC-025) ----------------------------------------------

  /** Opens the link dialog (AC-1, AC-2), prefilled from the selection or the
   * existing link the caret sits inside (AC-9). */
  function openLinkDialog() {
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    if (!root || !model || !editable.value) {
      return;
    }
    dismissLinkChip();
    const selection = currentSelection(root) ?? caretSelection(documentStart());
    const extent = isCollapsed(selection) ? linkAtCaret(model, selection.focus) : null;
    if (extent) {
      linkDialogSelection = extentSelection(extent);
      linkDialogInitial.value = {
        text: extent.text,
        url: extent.target.href ?? "",
        anchor: extent.target.anchor ?? "",
      };
    } else {
      linkDialogSelection = selection;
      linkDialogInitial.value = {
        text: isCollapsed(selection) ? "" : fragmentText(extractFragment(model, selection)),
        url: "",
        anchor: "",
      };
    }
    linkDialogOpen.value = true;
  }

  /** Applies the dialog's link (AC-3, AC-19, AC-20). */
  function handleLinkDialogSave(payload = { text: "", href: "", anchor: "" }) {
    const model = doc.value;
    const selection = linkDialogSelection;
    linkDialogOpen.value = false;
    linkDialogSelection = noSelection();
    if (!model || !selection) {
      return;
    }
    const anchor = String(payload.anchor ?? "");
    const href = anchor === "" ? sanitizeHref(String(payload.href ?? "")) : null;
    if (anchor === "" && href === null) {
      return;
    }
    const target = anchor !== "" ? { anchor } : { href: href ?? "" };
    const text = String(payload.text ?? "");
    const selectionText = isCollapsed(selection)
      ? ""
      : fragmentText(extractFragment(model, selection));
    const result =
      !isCollapsed(selection) && text === selectionText
        ? applyLink(model, selection, target)
        : insertLinkText(model, selection, text, target);
    applyEdit(result, selection, "insert link", false);
    (contentRoot.value as HTMLElement | null)?.focus();
  }

  function handleLinkDialogCancel() {
    linkDialogOpen.value = false;
    linkDialogSelection = noSelection();
    (contentRoot.value as HTMLElement | null)?.focus();
    void restoreCaretAfterRender();
  }

  function dismissLinkChip() {
    if (linkChip.value.visible) {
      linkChip.value = { visible: false, left: 0, top: 0, label: "", external: false };
    }
    linkChipExtent = noLinkExtent();
  }

  /** Shows the chip when a collapsed caret sits inside a link (AC-8). */
  async function refreshLinkChip() {
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    const caret = getCaret();
    if (!root || !model || !editable.value || !caret || getSelectionAnchor() || linkDialogOpen.value) {
      dismissLinkChip();
      return;
    }
    const at = caret;
    const extent = linkAtCaret(model, at);
    if (!extent) {
      dismissLinkChip();
      return;
    }
    await nextTick();
    if (getCaret() !== at || !contentRoot.value) {
      return;
    }
    const rect = caretClientRect(root, at);
    if (!rect) {
      dismissLinkChip();
      return;
    }
    linkChipExtent = extent;
    linkChip.value = {
      visible: true,
      left: Math.max(8, rect.left),
      top: rect.bottom + 8,
      label: linkTargetLabel(model, extent.target),
      external: extent.target.href !== undefined,
    };
  }

  /** Opens a link target: external URL in a new tab, or an in-document jump. */
  function openLinkTarget(target = noLinkTarget()) {
    if (!target) {
      return;
    }
    if (target.href !== undefined) {
      const href = sanitizeHref(target.href);
      if (href !== null) {
        window.open(href, "_blank", "noopener,noreferrer");
      }
      return;
    }
    jumpToAnchor(target.anchor ?? "");
  }

  /** Moves the caret to a bookmark or heading anchor and scrolls it into view. */
  function jumpToAnchor(anchor = "") {
    const model = doc.value;
    if (!model || anchor === "") {
      return;
    }
    const position = anchorCaret(model, anchor);
    if (!position) {
      return;
    }
    applyCaretMove(position, false);
    void scrollModelCaretIntoView(position);
  }

  function handleLinkChipOpen() {
    const extent = linkChipExtent;
    if (extent) {
      openLinkTarget(extent.target);
    }
  }

  function handleLinkChipEdit() {
    openLinkDialog();
  }

  /** Copies the chip's target to the clipboard (AC-10). */
  function handleLinkChipCopy() {
    const extent = linkChipExtent;
    if (!extent) {
      return;
    }
    const text = extent.target.href ?? `#${extent.target.anchor ?? ""}`;
    void navigator.clipboard?.writeText(text).catch(() => {});
  }

  /** Removes the link the caret sits in, keeping its text (AC-11, AC-33). */
  function handleLinkChipRemove() {
    const model = doc.value;
    const caret = getCaret();
    if (!model || !editable.value || !caret) {
      return;
    }
    const selection = caretSelection(caret);
    applyEdit(removeLink(model, selection), selection, "remove link", false);
    (contentRoot.value as HTMLElement | null)?.focus();
  }

  /** Autolinks after a completed space/Enter edit (AC-5, AC-6): a bare URL
   * ending where the word ended becomes a link, as its own undo step. */
  function maybeAutolink(result = noCommandResult(), afterEnter = false) {
    const model = doc.value;
    if (!model || !result) {
      return;
    }
    const focus = result.selection.focus;
    const cell = focus.cell;
    if (afterEnter && (focus.offset !== 0 || (cell ? cell.block === 0 : focus.block === 0))) {
      return;
    }
    const address = afterEnter
      ? cell
        ? {
            section: focus.section,
            block: focus.block,
            cell: { row: cell.row, col: cell.col, block: cell.block - 1 },
          }
        : { section: focus.section, block: focus.block - 1 }
      : cell
        ? { section: focus.section, block: focus.block, cell }
        : { section: focus.section, block: focus.block };
    const paragraph = paragraphAt(model, address);
    if (!paragraph) {
      return;
    }
    const endOffset = afterEnter ? blockLength(paragraph) : focus.offset - 1;
    if (endOffset <= 0) {
      return;
    }
    const linked = autolinkAt(model, address, endOffset);
    if (!linked) {
      return;
    }
    const keep = caretSelection(focus);
    applyEdit({ model: linked, selection: keep }, keep, "autolink", false);
  }

  // --- bookmarks (SPEC-025) -----------------------------------------------

  function openBookmarkDialog() {
    const root = contentRoot.value as HTMLElement | null;
    if (!root || !doc.value || !editable.value) {
      return;
    }
    bookmarkDialogSelection = currentSelection(root) ?? caretSelection(documentStart());
    bookmarkDialogOpen.value = true;
  }

  function handleBookmarkDialogSave(name = "") {
    const model = doc.value;
    const selection = bookmarkDialogSelection;
    bookmarkDialogOpen.value = false;
    bookmarkDialogSelection = noSelection();
    if (!model || !selection || name === "") {
      return;
    }
    if (validateBookmarkName(name, documentBookmarkNames()) !== "") {
      return;
    }
    applyEdit(insertBookmark(model, selection, name), selection, "insert bookmark", false);
    (contentRoot.value as HTMLElement | null)?.focus();
  }

  function handleBookmarkDialogCancel() {
    bookmarkDialogOpen.value = false;
    bookmarkDialogSelection = noSelection();
    (contentRoot.value as HTMLElement | null)?.focus();
    void restoreCaretAfterRender();
  }

  function handleBookmarkJump(name = "") {
    bookmarkManagerOpen.value = false;
    (contentRoot.value as HTMLElement | null)?.focus();
    jumpToAnchor(String(name));
  }

  function handleBookmarkRemove(name = "") {
    const model = doc.value;
    if (!model || !editable.value) {
      return;
    }
    const selection = currentSelection(contentRoot.value as HTMLElement | null) ?? caretSelection(documentStart());
    applyEdit(deleteBookmark(model, selection, String(name)), selection, "delete bookmark", false);
  }

  function closeBookmarkManager() {
    bookmarkManagerOpen.value = false;
    (contentRoot.value as HTMLElement | null)?.focus();
  }

  // --- table of contents (SPEC-033) ---------------------------------------

  /** A heading anchor's current displayed page number (SPEC-033). */
  function tocPageNumberFor(model = noDocumentPlaceholder(), anchor = ""): number | null {
    const position = anchorCaret(model, anchor);
    if (!position) {
      return null;
    }
    const pageIndex = pageOfCaret(position);
    return pageNumbers.value[pageIndex] ?? null;
  }

  function openTocDialog() {
    if (!doc.value || !editable.value || activeHeaderFooter.value) {
      return;
    }
    tocOptionsDialogOpen.value = true;
  }

  function handleTocInsert(options = defaultTocOptions()) {
    tocOptionsDialogOpen.value = false;
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    if (!root || !model || !editable.value) {
      return;
    }
    const selection = paragraphSelection(root);
    applyEdit(insertTableOfContents(model, selection, options, (anchor) => tocPageNumberFor(model, anchor)), selection, "insert table of contents");
    root.focus();
  }

  function handleTocCancel() {
    tocOptionsDialogOpen.value = false;
  }

  /** Deletes the selected TOC block (SPEC-033 AC-16). */
  function handleTocDelete() {
    const location = tocSelection.value;
    const model = doc.value;
    if (!location || !model || !editable.value) {
      return;
    }
    const start = { section: location.section, block: location.blockIndex, offset: 0 };
    const end = { section: location.section, block: location.blockIndex, offset: 1 };
    const selection = caretSelection(start);
    applyEdit(deleteRange(model, { anchor: start, focus: end }), selection, "delete table of contents");
  }

  /** DocTocBlock's single `action` event (SPEC-033 AC-8..13, AC-14/15). */
  function handleTocAction(payload = { type: "jump", anchor: "", section: 0, blockIndex: 0 }) {
    if (payload.type === "jump") {
      clearTocSelection();
      jumpToAnchor(payload.anchor ?? "");
    } else if (payload.type === "update-page-numbers") {
      handleTocUpdatePageNumbers(payload.section, payload.blockIndex);
    } else if (payload.type === "update-entire") {
      handleTocUpdateEntire(payload.section, payload.blockIndex);
    } else if (payload.type === "select") {
      clearCellSelection();
      clearImageSelection();
      tocSelection.value = { section: payload.section, blockIndex: payload.blockIndex };
    }
  }

  function handleTocUpdatePageNumbers(section = 0, blockIndex = 0) {
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    if (!root || !model || !editable.value) {
      return;
    }
    const selection = currentSelection(root) ?? caretSelection(getCaret() ?? documentStart());
    applyEdit(
      updateTocPageNumbersOnly(model, selection, section, blockIndex, (anchor) => tocPageNumberFor(model, anchor)),
      selection,
      "update page numbers",
    );
  }

  function handleTocUpdateEntire(section = 0, blockIndex = 0) {
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    if (!root || !model || !editable.value) {
      return;
    }
    const selection = currentSelection(root) ?? caretSelection(getCaret() ?? documentStart());
    applyEdit(
      updateTocEntirely(model, selection, section, blockIndex, (anchor) => tocPageNumberFor(model, anchor)),
      selection,
      "update table of contents",
    );
  }

  return {
    linkDialogOpen,
    linkDialogInitial,
    linkChip,
    bookmarkDialogOpen,
    bookmarkManagerOpen,
    tocSelection,
    tocOptionsDialogOpen,
    documentHeadings,
    documentBookmarkTargets,
    documentBookmarkNames,
    openLinkDialog,
    handleLinkDialogSave,
    handleLinkDialogCancel,
    dismissLinkChip,
    refreshLinkChip,
    openLinkTarget,
    jumpToAnchor,
    handleLinkChipOpen,
    handleLinkChipEdit,
    handleLinkChipCopy,
    handleLinkChipRemove,
    openBookmarkDialog,
    handleBookmarkDialogSave,
    handleBookmarkDialogCancel,
    handleBookmarkJump,
    handleBookmarkRemove,
    closeBookmarkManager,
    maybeAutolink,
    tocPageNumberFor,
    openTocDialog,
    handleTocInsert,
    handleTocCancel,
    handleTocDelete,
    handleTocAction,
    handleTocUpdatePageNumbers,
    handleTocUpdateEntire,
    clearTocSelection,
  };
}

function noDocumentPlaceholder(): DocumentModel {
  return { id: "", title: "", sections: [], defaultStyles: {} as DocumentModel["defaultStyles"] };
}
