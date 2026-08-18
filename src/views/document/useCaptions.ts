// Captions and the table of figures (SPEC-055) — the DocumentView wiring for
// Insert → Caption and Insert → Table of figures, the "delete its caption too?"
// prompt, the auto-caption setting, and the ToF block's update/jump actions.
// A composable in the SPEC-055 style: DocumentView spreads its EditorContext
// plus the object-selection refs and the shared anchor/jump helpers, and this
// owns all caption feature state so DocumentView gains only wiring.
import { computed, ref, type Ref } from "vue";
import type { EditorContext, HeaderFooterBridge } from "./context";
import type { DocumentModel, TofOptions } from "../../lib/documentModel";
import { isChart, isImage, isTof } from "../../lib/documentModel";
import type { CellRange } from "../../lib/core/tables";
import type { ImageLocation } from "../../lib/images";
import type { ChartLocation } from "../../lib/charts";
import type { CaptionOptions } from "../../lib/captions";
import {
  BUILT_IN_CAPTION_LABELS,
  captionForObject,
  defaultCaptionOptions,
  deleteObjectWithCaption,
  insertCaption,
  insertSurfaceCaption,
  isCaptionableObject,
  listCaptions,
} from "../../lib/captions";
import {
  availableTofLabels,
  defaultTofOptions,
  insertTableOfFigures,
  noTofSelection,
  updateTofEntirely,
  updateTofPageNumbersOnly,
} from "../../lib/tof";
import { caretSelection, deleteRange, documentStart, withHeaderFooterModel } from "../../lib/commands";
import type { CaretPosition, ParagraphAddress } from "../../lib/textEditing";
import {
  loadCaptionSettings,
  saveCaptionSettings,
  type CaptionSettings,
} from "../../lib/core/captionSettings";

export interface CaptionsDeps extends EditorContext, HeaderFooterBridge {
  /** The active band's editable root (focus restore after a band insert). */
  hfRootEl: () => HTMLElement | null;
  /** The currently selected image (SPEC-024), if any. */
  imageSelection: Ref<ImageLocation | null>;
  /** The currently selected chart (SPEC-052), if any. */
  chartSelection: Ref<ChartLocation | null>;
  /** A rectangular table cell selection (SPEC-022), if any. */
  cellSelection: Ref<CellRange | null>;
  /** A caption anchor's current page number — DocumentView passes useLinks's
   * tocPageNumberFor, whose anchorCaret already resolves caption anchors. */
  tofPageNumberFor: (model: DocumentModel, anchor: string) => number | null;
  /** Scroll to an internal anchor (a caption entry click). */
  jumpToAnchor: (anchor: string) => void;
  /** The reactive caret mirror (DocumentView's caretPosition) — so
   * captionTargetSelected re-evaluates when the caret enters/leaves a table
   * cell (a plain getCaret() read would make the computed stale). */
  caretPosition: Ref<CaretPosition | null>;
  /** The object composables' own delete/insert paths, wrapped here so a
   * captioned object's deletion prompts first (AC-20) and a freshly inserted
   * object auto-captions (AC-19). */
  handleImageDelete: () => void;
  handleChartDelete: () => void;
  insertImageFromFile: (file?: File) => Promise<void>;
  handleInsertTable: (size?: { rows: number; cols: number }) => void;
}

export function useCaptions(deps: CaptionsDeps) {
  const {
    doc,
    contentRoot,
    editable,
    currentSelection,
    getCaret,
    applyEdit,
    addToast,
    clearCellSelection,
  } = deps;

  const captionDialogOpen = ref(false);
  const tofDialogOpen = ref(false);
  const tofSelection = ref(noTofSelection());
  const captionSettings = ref<CaptionSettings>(loadCaptionSettings());

  // The delete-caption prompt (AC-20/21): the object awaiting a decision and
  // the caller's own object-delete, run when the user chooses "keep caption".
  const deleteCaptionPromptOpen = ref(false);
  const deleteCaptionKind = ref("object");
  let pendingObject: ParagraphAddress | null = null;
  let pendingCaption: ParagraphAddress | null = null;
  let pendingDelete: (() => void) | null = null;

  /** The object (image/chart/table) the caption commands act on, resolved from
   * whatever is selected: a selected image or chart, else the table the caret
   * (or a cell-range drag) sits in. Returns null when nothing captionable is
   * selected. This is the fix for the prior round's AC-1/AC-2 failure — a
   * table is now recognised from a plain caret in a cell, not only a
   * multi-cell drag. */
  function selectedObject(): ParagraphAddress | null {
    const model = doc.value;
    if (!model) {
      return null;
    }
    if (deps.imageSelection.value) {
      return { section: deps.imageSelection.value.section, block: deps.imageSelection.value.block };
    }
    if (deps.chartSelection.value) {
      return { section: deps.chartSelection.value.section, block: deps.chartSelection.value.block };
    }
    const range = deps.cellSelection.value;
    if (range) {
      const block = model.sections[range.section]?.blocks[range.block];
      if (block?.type === "table") {
        return { section: range.section, block: range.block };
      }
    }
    const caret = deps.caretPosition.value ?? getCaret();
    if (caret?.cell) {
      const block = model.sections[caret.section]?.blocks[caret.block];
      if (block?.type === "table") {
        return { section: caret.section, block: caret.block };
      }
    }
    return null;
  }

  /** The active header/footer band as its synthetic single-section model
   * (SPEC-029's editing shape), or null during body editing. Captions are
   * constructible inside a band (AC-37) because SPEC-029 AC-5/AC-6 already
   * put tables and images there. */
  function bandModel(): DocumentModel | null {
    const active = deps.activeHeaderFooter.value;
    const model = doc.value;
    return active && model
      ? withHeaderFooterModel(model, active.sectionIndex, active.part, active.kind)
      : null;
  }

  /** The captionable object at the band's live selection: the table whose
   * cell the caret sits in, else the image/chart block at or adjacent to the
   * caret — the band has no click-selection refs, so adjacency stands in for
   * the body's imageSelection/chartSelection. */
  function bandSelectedObject(band: DocumentModel): ParagraphAddress | null {
    const caret = deps.currentHfSelection()?.focus;
    if (!caret) {
      return null;
    }
    const blocks = band.sections[caret.section]?.blocks ?? [];
    if (caret.cell && blocks[caret.block]?.type === "table") {
      return { section: caret.section, block: caret.block };
    }
    for (const blockIndex of [caret.block, caret.block - 1, caret.block + 1]) {
      const block = blocks[blockIndex];
      if (isImage(block) || isChart(block)) {
        return { section: caret.section, block: blockIndex };
      }
    }
    return null;
  }

  /** Whether a captionable object is selected — enables the menu item and the
   * context-menu action (AC-1). In a header/footer band the caret is
   * deliberately not reactive (see useHeadersFooters), so the item enables
   * whenever the band holds a captionable object and openCaptionDialog
   * re-resolves the exact target from the live selection at click time. */
  const captionTargetSelected = computed(() => {
    const band = bandModel();
    if (band) {
      return band.sections[0].blocks.some((block) => isCaptionableObject(block));
    }
    return selectedObject() !== null;
  });

  /** The labels the caption dialog offers: the built-ins plus any custom label
   * already used in the document. */
  function dialogLabels(model: DocumentModel): string[] {
    const custom = listCaptions(model)
      .map((c) => c.label)
      .filter((label) => !BUILT_IN_CAPTION_LABELS.some((b) => b.toLowerCase() === label.toLowerCase()));
    return [...BUILT_IN_CAPTION_LABELS, ...Array.from(new Set(custom))];
  }

  const captionLabels = computed(() => (doc.value ? dialogLabels(doc.value) : BUILT_IN_CAPTION_LABELS));

  function openCaptionDialog() {
    if (!doc.value || !editable.value) {
      return;
    }
    const band = bandModel();
    if (band ? !bandSelectedObject(band) : !selectedObject()) {
      addToast("Select an image, chart or table first");
      return;
    }
    captionDialogOpen.value = true;
  }

  function handleCaptionInsert(options: CaptionOptions = defaultCaptionOptions()) {
    captionDialogOpen.value = false;
    if (!editable.value) {
      return;
    }
    const band = bandModel();
    if (band) {
      // A band caption commits through the SPEC-029 synthetic-model path and
      // is never renumbered (AC-37, AC-38, §Numbering): it keeps the cached
      // number it is created with, and the body counter skips it.
      const object = bandSelectedObject(band);
      if (!object) {
        return;
      }
      const result = insertSurfaceCaption(band, object, options);
      deps.applyHeaderFooterEdit(result.model, result.selection, "insert caption");
      deps.hfRootEl()?.focus();
      return;
    }
    const model = doc.value;
    const object = selectedObject();
    const root = contentRoot.value as HTMLElement | null;
    if (!model || !object) {
      return;
    }
    const before = currentSelection(root) ?? caretSelection(getCaret() ?? documentStart());
    applyEdit(insertCaption(model, object, options), before, "insert caption");
    root?.focus();
  }

  function handleCaptionCancel() {
    captionDialogOpen.value = false;
  }

  /** Auto-captions a freshly inserted object when the setting is on (AC-18,
   * AC-19). `objectKind` is "image" or "table"; charts are excluded by design. */
  function maybeAutoCaption(object: ParagraphAddress, objectKind: "image" | "table") {
    const model = doc.value;
    if (!model || !editable.value) {
      return;
    }
    if (objectKind === "image" && !captionSettings.value.autoImages) {
      return;
    }
    if (objectKind === "table" && !captionSettings.value.autoTables) {
      return;
    }
    const label = objectKind === "table" ? "Table" : "Figure";
    const before = caretSelection(getCaret() ?? documentStart());
    applyEdit(insertCaption(model, object, defaultCaptionOptions(label)), before, "auto caption");
  }

  /** Auto-captions the object just inserted at/near the caret (AC-18, AC-19),
   * when the matching setting is on. Called after an image or table insert;
   * scans the caret's block and its neighbours so it works whether the insert
   * left the caret inside the object or just after it, and never
   * double-captions an object that already has one. */
  function autoCaptionNearCaret(objectKind: "image" | "table") {
    // A band insert (SPEC-029 AC-5/AC-6) is never auto-captioned: the body
    // caret this scan reads is stale while a band is active, and band
    // captions sit outside the numbering anyway (§Numbering).
    if (deps.activeHeaderFooter.value) {
      return;
    }
    const model = doc.value;
    const caret = getCaret();
    if (!model || !caret) {
      return;
    }
    const blocks = model.sections[caret.section]?.blocks ?? [];
    for (const blockIndex of [caret.block, caret.block - 1, caret.block + 1]) {
      const block = blocks[blockIndex];
      if (block?.type === objectKind) {
        const object = { section: caret.section, block: blockIndex };
        if (!captionForObject(model, object)) {
          maybeAutoCaption(object, objectKind);
        }
        return;
      }
    }
  }

  function toggleAutoImages() {
    const next = { ...captionSettings.value, autoImages: !captionSettings.value.autoImages };
    captionSettings.value = next;
    saveCaptionSettings(next);
  }

  function toggleAutoTables() {
    const next = { ...captionSettings.value, autoTables: !captionSettings.value.autoTables };
    captionSettings.value = next;
    saveCaptionSettings(next);
  }

  // --- delete-with-caption prompt (AC-20, AC-21, AC-22) -------------------

  /** Runs `performDelete` for an object, first prompting whether to delete its
   * bound caption too when it has one (AC-20). No prompt for an uncaptioned
   * object (AC-22) — `performDelete` runs immediately. */
  function requestObjectDelete(object: ParagraphAddress, objectKind: string, performDelete: () => void) {
    const model = doc.value;
    const caption = model ? captionForObject(model, object) : null;
    if (!model || !caption) {
      performDelete();
      return;
    }
    pendingObject = object;
    pendingCaption = caption;
    pendingDelete = performDelete;
    deleteCaptionKind.value = objectKind;
    deleteCaptionPromptOpen.value = true;
  }

  function confirmDeleteBoth() {
    const model = doc.value;
    deleteCaptionPromptOpen.value = false;
    if (model && pendingObject && pendingCaption && editable.value) {
      const before = caretSelection(getCaret() ?? documentStart());
      applyEdit(deleteObjectWithCaption(model, pendingObject, pendingCaption), before, "delete object and caption");
    }
    resetDeletePrompt();
  }

  function confirmKeepCaption() {
    deleteCaptionPromptOpen.value = false;
    pendingDelete?.();
    resetDeletePrompt();
  }

  function cancelDeletePrompt() {
    deleteCaptionPromptOpen.value = false;
    resetDeletePrompt();
  }

  function resetDeletePrompt() {
    pendingObject = null;
    pendingCaption = null;
    pendingDelete = null;
  }

  /** Deleting a selected image/chart prompts to remove a bound caption too
   * (AC-20), falling through to the plain delete when uncaptioned. */
  function deleteImageWithCaptionPrompt() {
    const selected = deps.imageSelection.value;
    if (selected) {
      requestObjectDelete({ section: selected.section, block: selected.block }, "image", deps.handleImageDelete);
    } else {
      deps.handleImageDelete();
    }
  }

  function deleteChartWithCaptionPrompt() {
    const selected = deps.chartSelection.value;
    if (selected) {
      requestObjectDelete({ section: selected.section, block: selected.block }, "chart", deps.handleChartDelete);
    } else {
      deps.handleChartDelete();
    }
  }

  /** Insert an image from a file, then auto-caption it if the setting is on
   * (AC-19). */
  async function insertImageFileWithAutoCaption(file = new File([], "")) {
    await deps.insertImageFromFile(file);
    autoCaptionNearCaret("image");
  }

  /** Insert a table, then auto-caption it if the setting is on (AC-19). */
  function insertTableWithAutoCaption(size = { rows: 1, cols: 1 }) {
    deps.handleInsertTable(size);
    autoCaptionNearCaret("table");
  }

  // --- table of figures --------------------------------------------------

  const tofLabels = computed(() => (doc.value ? availableTofLabels(doc.value) : ["Figure"]));

  function openTableOfFiguresDialog() {
    if (!doc.value || !editable.value) {
      return;
    }
    tofDialogOpen.value = true;
  }

  function handleTofInsert(payload: { label: string; options: TofOptions } = { label: "Figure", options: defaultTofOptions() }) {
    tofDialogOpen.value = false;
    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    if (!model || !root || !editable.value) {
      return;
    }
    const selection = currentSelection(root) ?? caretSelection(getCaret() ?? documentStart());
    applyEdit(
      insertTableOfFigures(model, selection, payload.label, payload.options, (anchor) => deps.tofPageNumberFor(model, anchor)),
      selection,
      "insert table of figures",
    );
    root.focus();
  }

  function handleTofCancel() {
    tofDialogOpen.value = false;
  }

  function handleTofAction(payload = { type: "jump", anchor: "", section: 0, blockIndex: 0 }) {
    if (payload.type === "jump") {
      clearTofSelection();
      deps.jumpToAnchor(payload.anchor ?? "");
    } else if (payload.type === "update-page-numbers") {
      updateTof(payload.section, payload.blockIndex, false);
    } else if (payload.type === "update-entire") {
      updateTof(payload.section, payload.blockIndex, true);
    } else if (payload.type === "select") {
      clearCellSelection();
      tofSelection.value = { section: payload.section, blockIndex: payload.blockIndex };
    }
  }

  function updateTof(section: number, blockIndex: number, entire: boolean) {
    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    if (!model || !editable.value) {
      return;
    }
    const selection = currentSelection(root) ?? caretSelection(getCaret() ?? documentStart());
    const lookup = (anchor: string) => deps.tofPageNumberFor(model, anchor);
    const result = entire
      ? updateTofEntirely(model, selection, section, blockIndex, lookup)
      : updateTofPageNumbersOnly(model, selection, section, blockIndex, lookup);
    applyEdit(result, selection, entire ? "update table of figures" : "update figure page numbers");
  }

  function clearTofSelection() {
    tofSelection.value = null;
  }

  /** Deletes the selected table-of-figures block; captions are untouched
   * (AC-35). */
  function handleTofDelete(): boolean {
    const location = tofSelection.value;
    const model = doc.value;
    if (!location || !model || !editable.value) {
      return false;
    }
    const block = model.sections[location.section]?.blocks[location.blockIndex];
    if (!isTof(block)) {
      return false;
    }
    const start = { section: location.section, block: location.blockIndex, offset: 0 };
    const end = { section: location.section, block: location.blockIndex, offset: 1 };
    applyEdit(deleteRange(model, { anchor: start, focus: end }), caretSelection(start), "delete table of figures");
    clearTofSelection();
    return true;
  }

  return {
    captionDialogOpen,
    captionLabels,
    captionTargetSelected,
    openCaptionDialog,
    handleCaptionInsert,
    handleCaptionCancel,
    captionSettings,
    toggleAutoImages,
    toggleAutoTables,
    requestObjectDelete,
    deleteCaptionPromptOpen,
    deleteCaptionKind,
    confirmDeleteBoth,
    confirmKeepCaption,
    cancelDeletePrompt,
    deleteImageWithCaptionPrompt,
    deleteChartWithCaptionPrompt,
    insertImageFileWithAutoCaption,
    insertTableWithAutoCaption,
    tofDialogOpen,
    tofLabels,
    tofSelection,
    openTableOfFiguresDialog,
    handleTofInsert,
    handleTofCancel,
    handleTofAction,
    handleTofDelete,
    clearTofSelection,
  };
}
