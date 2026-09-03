// The style picker's actions (SPEC-019) and the custom styles manager
// (SPEC-042): applying a paragraph style from the picker or its shortcuts,
// "update to match"/"reset style", the Styles pane, the New/Modify style
// editor, delete confirmation, applying styles of every family from the
// pane, and the import-styles-from-another-document flow (AC-19..24). The
// caret's style *readouts* (styleState / characterStyleState) stay in
// DocumentView's updateCaretStatus, which recomputes them with the rest of
// the toolbar state; this composable owns what the pane and dialogs render
// and the commands they fire.
import { computed, ref } from "vue";
import {
  applyCharacterStyle,
  applyParagraphStyle,
  applyTableStyle,
  caretFormatting,
  createStyle,
  deleteStyle,
  modifyStyle,
  resetParagraphStyle,
  updateStyleFromSelection,
} from "../../lib/commands";
import {
  importStyles,
  noStyleImportPlan,
  noStyleImportResolutionMap,
  noStyleImportSummary,
  planStyleImport,
  readStylesFromFile,
  styleImportConflictRows,
  styleImportResolutionsFromMap,
} from "../../lib/customStyles";
import { noNamedStyles } from "../../lib/documentModel";
import {
  allStyles,
  isStyleInUse,
  noStyleEntries,
  noStylesPaneEntries,
  resolveNamedStyle,
  styleById,
  stylePickerEntries,
  stylesPaneEntries,
  styleTypeOf,
} from "../../lib/styles";
import { paragraphAtCaret } from "../../lib/textEditing";
import { normaliseParagraphBorders, paintedShadingFill } from "../../lib/paragraphBorders";
import type { ModelSelection, ParagraphCommand, StyleFormInput } from "../../lib/commands";
import type { PendingFormat } from "../../lib/documentModel";
import type { CaretPosition } from "../../lib/textEditing";
import type { EditorContext } from "./context";

export interface StyleManagerDeps extends EditorContext {
  /** DocumentView's paragraphSelection always resolves to a selection. */
  paragraphSelection: (root?: HTMLElement | null) => ModelSelection;
  /** The view's shared paragraph-command dispatcher: runs one command
   * against the current selection (routing into an active header/footer or
   * note band when one is live) and applies its result under `label`;
   * `refocus` returns focus to the canvas. */
  applyParagraphCommand: (
    command: ParagraphCommand | null,
    label: string,
    refocus?: boolean,
  ) => void;
  /** Live write of the view's pending collapsed-caret format (SPEC-012):
   * applying a style to an empty paragraph has no run to carry the style's
   * run formatting, so it becomes the pending format the next typed text
   * takes. */
  setPendingFormat: (next: PendingFormat | null) => void;
  updateCaretStatus: (range?: { start: CaretPosition; end: CaretPosition }) => void;
  /** Live read of the view's reactive caret mirror (`caretPosition`) — a
   * getter over a ref, so a computed reading it re-runs when the caret
   * moves. */
  getCaretPosition: () => CaretPosition | null;
}

export function useStyleManager(deps: StyleManagerDeps) {
  const {
    doc,
    contentRoot,
    editable,
    applyEdit,
    paragraphSelection,
    applyParagraphCommand,
    setPendingFormat,
    updateCaretStatus,
    getCaretPosition,
  } = deps;

  // --- paragraph styles (SPEC-019) --------------------------------------------
  // The style picker's menu entries — the document's own styles, or the
  // built-in set for a blank document (AC-2, AC-8). doc.value is replaced on
  // every edit, so the list follows catalogue changes (AC-13).
  const styleEntries = computed(() => (doc.value ? stylePickerEntries(doc.value) : noStyleEntries()));
  // --- custom styles manager (SPEC-042) ---------------------------------------
  // The Styles pane/dialogs' own state.
  const stylesPaneOpen = ref(false);
  const stylesPaneEntriesList = computed(() => (doc.value ? stylesPaneEntries(doc.value) : noStylesPaneEntries()));
  const styleEditorOpen = ref(false);
  const styleEditorMode = ref("new");
  const styleEditorType = ref("paragraph");
  const styleEditorTargetId = ref("");
  const styleDeleteTargetId = ref("");
  const styleDeleteTargetInUse = ref(false);
  const importDialogOpen = ref(false);
  const importPhase = ref("pick");
  const importFileName = ref("");
  const importError = ref("");
  const importSourceStyles = ref(noNamedStyles());
  const importPlanState = ref(noStyleImportPlan());
  const importResolutions = ref(noStyleImportResolutionMap());
  const importSummaryState = ref(noStyleImportSummary());

  // --- paragraph styles (SPEC-019) --------------------------------------------

  /** Applies a style to every paragraph in the selection (AC-3) — the picker's
   * click and the Ctrl/Cmd+Alt+0..6 shortcuts (AC-5, AC-6) both land here. An
   * empty caret paragraph has no run to carry the style's run formatting, so
   * it becomes the pending format the next typed text takes. */
  function handleApplyStyle(styleId: string): void {
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    if (styleId === "" || !root || !model || !editable.value) {
      return;
    }
    const selection = paragraphSelection(root);
    applyEdit(applyParagraphStyle(model, selection, styleId), selection, "paragraph style");
    const after = doc.value;
    const block = after ? paragraphAtCaret(after, selection.focus) : undefined;
    if (after && block && block.runs.length === 0) {
      setPendingFormat({ ...resolveNamedStyle(after, styleId).runFormatting });
      updateCaretStatus({ start: selection.anchor, end: selection.focus });
    }
    root.focus();
  }

  /** "Update <style> to match" from a picker entry's context menu (AC-12). */
  function handleUpdateStyle(styleId: string): void {
    if (styleId === "") {
      return;
    }
    applyParagraphCommand(
      (model, selection) => updateStyleFromSelection(model, selection, styleId),
      "update style",
    );
  }

  /** "Reset style": removes direct formatting from the selection (AC-14). */
  function handleResetStyle(): void {
    applyParagraphCommand(
      (model, selection) => resetParagraphStyle(model, selection),
      "reset style",
    );
  }

  // --- custom styles manager (SPEC-042) ---------------------------------------

  function toggleStylesPane(): void {
    stylesPaneOpen.value = !stylesPaneOpen.value;
  }

  function closeStylesPane(): void {
    stylesPaneOpen.value = false;
  }

  /** "New style" (AC-3): seeded from the caret's current formatting, and — for
   * a paragraph/list style — its current paragraph style as the "based on"
   * starting point. */
  function openNewStyle(type: string): void {
    styleEditorMode.value = "new";
    styleEditorType.value = type;
    styleEditorTargetId.value = "";
    styleEditorOpen.value = true;
  }

  /** "Modify style" (AC-4). */
  function openModifyStyle(styleId: string): void {
    const model = doc.value;
    const target = model ? styleById(model, styleId) : undefined;
    styleEditorMode.value = "modify";
    styleEditorType.value = target ? styleTypeOf(target) : "paragraph";
    styleEditorTargetId.value = styleId;
    styleEditorOpen.value = true;
  }

  function closeStyleEditor(): void {
    styleEditorOpen.value = false;
  }

  /** The style being modified, or an empty style for "new" (fields default to
   * the caret's own formatting via styleEditorInitialProps below). */
  const styleEditorInitialProps = computed(() => {
    const model = doc.value;
    if (styleEditorMode.value === "modify") {
      const target = model ? styleById(model, styleEditorTargetId.value) : undefined;
      const border = target?.cellBorders?.top;
      return {
        initialName: target?.name ?? "",
        initialType: target ? styleTypeOf(target) : "paragraph",
        initialBasedOn: target?.basedOn ?? "",
        initialNext: target?.next ?? "",
        initialFontFamily: target?.runFormatting?.fontFamily ?? "",
        initialFontSizePt: target?.runFormatting?.fontSizePt,
        initialBold: target?.runFormatting?.bold === true,
        initialItalic: target?.runFormatting?.italic === true,
        initialUnderline: target?.runFormatting?.underline === true,
        initialColor: target?.runFormatting?.color ?? "",
        initialAlignment: target?.paragraphFormatting?.alignment ?? "",
        initialListPresetKey: "decimal",
        initialCellShadingFill: target?.cellShadingFill ?? "",
        initialCellBorderStyle: border?.style ?? "",
        initialCellBorderWidthPt: border?.widthPt ?? 1,
        initialCellBorderColor: border?.color ?? "#000000",
        // SPEC-074 AC-18: the style's own paragraph borders and shading.
        // Seeded through the paint helpers, so a value that is present but
        // paints nothing — AC-17's explicit "off", which a .docx style can
        // carry as an all-nil w:pBdr or w:shd w:fill="auto" — comes up as
        // plain "no borders"/"no fill" rather than as a keyword in a colour
        // input. The editor has no separate control for overriding a parent.
        initialBorders: normaliseParagraphBorders(target?.paragraphFormatting?.borders),
        initialShadingFill: paintedShadingFill(target?.paragraphFormatting) ?? "",
      };
    }
    const caretNow = getCaretPosition();
    const caretRun = model && caretNow ? caretFormatting(model, caretNow) : {};
    const caretParagraph = model && caretNow ? paragraphAtCaret(model, caretNow) : undefined;
    return {
      initialName: "",
      initialType: styleEditorType.value,
      initialBasedOn: caretParagraph?.styleId ?? "",
      initialNext: "",
      initialFontFamily: caretRun.fontFamily ?? "",
      initialFontSizePt: caretRun.fontSizePt,
      initialBold: caretRun.bold === true,
      initialItalic: caretRun.italic === true,
      initialUnderline: caretRun.underline === true,
      initialColor: caretRun.color ?? "",
      initialAlignment: caretParagraph?.formatting?.alignment ?? "",
      initialListPresetKey: "decimal",
      initialCellShadingFill: "",
      initialCellBorderStyle: "",
      initialCellBorderWidthPt: 1,
      initialCellBorderColor: "#000000",
      // A new style is seeded from the caret's paragraph, like every other
      // field above (SPEC-074 AC-18) — and through the same paint helpers, so
      // a paragraph carrying AC-17's explicit "off" seeds the style with no
      // borders and no fill rather than with the override itself.
      initialBorders: normaliseParagraphBorders(caretParagraph?.formatting?.borders),
      initialShadingFill: paintedShadingFill(caretParagraph?.formatting) ?? "",
    };
  });

  const styleEditorBasedOnOptions = computed(() => {
    const model = doc.value;
    if (!model) {
      return [];
    }
    return allStyles(model)
      .filter((entry) => entry.name !== undefined && entry.id !== styleEditorTargetId.value)
      .map((entry) => ({ id: entry.id, name: entry.name ?? entry.id }));
  });

  const styleEditorExistingNames = computed(() => {
    const model = doc.value;
    if (!model) {
      return [];
    }
    return allStyles(model)
      .filter((entry) => entry.name !== undefined && entry.id !== styleEditorTargetId.value)
      .map((entry) => (entry.name ?? entry.id).trim().toLowerCase());
  });

  function saveStyleEditor(payload: StyleFormInput): void {
    if (styleEditorMode.value === "new") {
      applyParagraphCommand((model, selection) => createStyle(model, selection, payload), "new style");
    } else {
      applyParagraphCommand(
        (model, selection) => modifyStyle(model, selection, styleEditorTargetId.value, payload),
        "modify style",
      );
    }
    styleEditorOpen.value = false;
  }

  /** "Delete style" (AC-7): the pane's own Delete button opens the confirm
   * dialog directly — the in-use warning (AC-8) is computed from the live
   * document here, not left to the dialog to guess. */
  function openDeleteStyle(styleId: string): void {
    const model = doc.value;
    styleDeleteTargetId.value = styleId;
    styleDeleteTargetInUse.value = model ? isStyleInUse(model, styleId) : false;
  }

  function cancelDeleteStyle(): void {
    styleDeleteTargetId.value = "";
  }

  function confirmDeleteStyle(): void {
    applyParagraphCommand(
      (model, selection) => deleteStyle(model, selection, styleDeleteTargetId.value),
      "delete style",
    );
    styleDeleteTargetId.value = "";
  }

  const styleDeleteTargetName = computed(() => {
    const model = doc.value;
    const target = model ? styleById(model, styleDeleteTargetId.value) : undefined;
    return target?.name ?? styleDeleteTargetId.value;
  });

  /** Applying a style from the pane (AC-9): character/table styles use their
   * own commands; paragraph/list styles reuse the picker's own apply path
   * (which also links a list style's numbering, SPEC-032's styleLinkFor). */
  function applyStyleFromPane(payload: { id: string; type: string }): void {
    if (payload.type === "character") {
      applyParagraphCommand(
        (model, selection) => applyCharacterStyle(model, selection, payload.id),
        "character style",
      );
    } else if (payload.type === "table") {
      applyParagraphCommand((model, selection) => applyTableStyle(model, selection, payload.id), "table style");
    } else {
      handleApplyStyle(payload.id);
    }
  }

  // --- import styles from another document (SPEC-042 AC-19..24) --------------

  function openImportDialog(): void {
    importDialogOpen.value = true;
    importPhase.value = "pick";
    importFileName.value = "";
    importError.value = "";
    importSourceStyles.value = noNamedStyles();
    importPlanState.value = noStyleImportPlan();
    importResolutions.value = noStyleImportResolutionMap();
    importSummaryState.value = noStyleImportSummary();
  }

  function closeImportDialog(): void {
    importDialogOpen.value = false;
  }

  const importConflictRows = computed(() =>
    importPlanState.value ? styleImportConflictRows(importPlanState.value.conflicts) : [],
  );
  const importNonConflictingCount = computed(() => importPlanState.value?.nonConflicting.length ?? 0);

  /** Reads the chosen file's styles and plans the merge (AC-19, AC-20, AC-21). */
  async function handleImportChooseFile(file: File): Promise<void> {
    const model = doc.value;
    if (!model) {
      return;
    }
    importError.value = "";
    try {
      const styles = await readStylesFromFile(file);
      importFileName.value = file.name;
      importSourceStyles.value = styles;
      importPlanState.value = planStyleImport(model, styles);
      importResolutions.value = noStyleImportResolutionMap();
      importPhase.value = "plan";
    } catch (error) {
      importError.value = error instanceof Error ? error.message : "Could not read styles from that file";
    }
  }

  function setImportResolution(payload: { sourceId: string; resolution: string }): void {
    importResolutions.value = { ...importResolutions.value, [payload.sourceId]: payload.resolution };
  }

  /** Applies the plan (AC-22, AC-23) and shows the summary (AC-24). */
  function confirmImportStyles(): void {
    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    if (!model || !root) {
      return;
    }
    const selection = paragraphSelection(root);
    const resolutions = styleImportResolutionsFromMap(importResolutions.value);
    const result = importStyles(model, importSourceStyles.value, resolutions);
    applyEdit({ model: result.model, selection }, selection, "import styles");
    importSummaryState.value = result.summary;
    importPhase.value = "summary";
  }

  return {
    styleEntries,
    stylesPaneOpen,
    stylesPaneEntriesList,
    styleEditorOpen,
    styleEditorMode,
    styleEditorType,
    styleEditorTargetId,
    styleDeleteTargetId,
    styleDeleteTargetInUse,
    importDialogOpen,
    importPhase,
    importFileName,
    importError,
    importSourceStyles,
    importPlanState,
    importResolutions,
    importSummaryState,
    styleEditorInitialProps,
    styleEditorBasedOnOptions,
    styleEditorExistingNames,
    styleDeleteTargetName,
    importConflictRows,
    importNonConflictingCount,
    handleApplyStyle,
    handleUpdateStyle,
    handleResetStyle,
    toggleStylesPane,
    closeStylesPane,
    openNewStyle,
    openModifyStyle,
    closeStyleEditor,
    saveStyleEditor,
    openDeleteStyle,
    cancelDeleteStyle,
    confirmDeleteStyle,
    applyStyleFromPane,
    openImportDialog,
    closeImportDialog,
    handleImportChooseFile,
    setImportResolution,
    confirmImportStyles,
  };
}
