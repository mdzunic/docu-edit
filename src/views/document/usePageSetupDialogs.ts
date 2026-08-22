// The Page setup dialog (SPEC-027 AC-1..7), the Format → Columns dialog
// (SPEC-030), the "Show/hide ruler" view toggle, and the ruler's
// caret-section geometry (page width, margins and column boundaries). With
// per-section page setup (SPEC-028 AC-20) both dialogs prefill from — and
// the ruler follows — the section under the caret, read live through
// getCaretSectionIndex. The ruler's marker-drag handler (handleRulerUpdate)
// stays in DocumentView: it dispatches indent/margin/column/tab-stop
// commands through the view's applyParagraphCommand chokepoint.
import { computed, ref, triggerRef } from "vue";
import { applyColumns, asColumnsScope } from "../../lib/columns";
import { asPageSetupScope, setPageSetup } from "../../lib/commands";
import { asMeasurementUnit } from "../../lib/core/units";
import {
  resolveColumnWidthsCm,
  sectionColumns,
  sectionContentWidthCm,
  sectionMargins,
  sectionPageSize,
} from "../../lib/documentModel";
import { touchDocument } from "../../lib/documentStore";
import type { ModelSelection } from "../../lib/commands";
import type { EditorContext } from "./context";

export interface PageSetupDialogsDeps extends EditorContext {
  /** DocumentView's paragraphSelection always resolves to a selection. */
  paragraphSelection: (root?: HTMLElement | null) => ModelSelection;
  /** Live read of the view's reactive caret-section mirror
   * (`caretSectionIndex`) — a getter over a ref, so a computed reading it
   * re-runs when the caret crosses into another section. */
  getCaretSectionIndex: () => number;
}

export function usePageSetupDialogs(deps: PageSetupDialogsDeps) {
  const { doc, contentRoot, editable, applyEdit, paragraphSelection, getCaretSectionIndex } = deps;

  // The Page setup dialog (SPEC-027 AC-1..7): open flag plus the prefill
  // taken from the document's own page geometry/unit at the moment it opens.
  const pageSetupOpen = ref(false);
  const pageSetupInitial = ref({
    pageSize: { widthCm: 21, heightCm: 29.7 },
    margins: { topCm: 2.54, rightCm: 2.54, bottomCm: 2.54, leftCm: 2.54 },
    measurementUnit: "cm",
    applyTo: "wholeDocument",
  });
  // The Format → Columns dialog (SPEC-030): open flag plus the prefill taken
  // from the caret's section columns and content width at the moment it opens.
  const columnsDialogOpen = ref(false);
  const columnsDialogInitial = ref({
    columns: { count: 1, equalWidth: true, gapCm: 1.27, separator: false },
    contentWidthCm: 16,
    applyTo: "thisSection",
  });

  // --- page setup and ruler visibility (SPEC-027) -----------------------

  /** Opens the Page setup dialog prefilled from the document's own page
   * geometry and measurement-unit preference (AC-8; absent unit reads as
   * "cm", createDefaultStyles's own convention). */
  function openPageSetup() {
    const model = doc.value;
    if (!model) {
      return;
    }
    // Prefilled from the caret's section (SPEC-028 AC-20): with per-section
    // page setup, "what the dialog shows" is the geometry under the caret.
    const sectionIndex = Math.min(getCaretSectionIndex(), model.sections.length - 1);
    pageSetupInitial.value = {
      pageSize: { ...sectionPageSize(model, sectionIndex) },
      margins: { ...sectionMargins(model, sectionIndex) },
      measurementUnit: model.measurementUnit ?? "cm",
      applyTo: "wholeDocument",
    };
    pageSetupOpen.value = true;
  }

  /** The dialog's OK (AC-5): one combined edit — page size, margins and the
   * measurement-unit preference — applied as a single undo step, which
   * re-renders the page and re-flows text immediately (doc.value's
   * replacement is what drives pageStyle/marginStyle's recompute). */
  function handlePageSetupApply(
    payload = {
      pageSize: { widthCm: 21, heightCm: 29.7 },
      margins: { topCm: 2.54, rightCm: 2.54, bottomCm: 2.54, leftCm: 2.54 },
      measurementUnit: "cm",
      applyTo: "wholeDocument",
    },
  ) {
    pageSetupOpen.value = false;
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    if (!root || !model || !editable.value) {
      return;
    }
    const selection = paragraphSelection(root);
    applyEdit(
      setPageSetup(
        model,
        selection,
        {
          pageSize: payload.pageSize,
          margins: payload.margins,
          measurementUnit: asMeasurementUnit(payload.measurementUnit),
        },
        asPageSetupScope(payload.applyTo ?? "wholeDocument"),
      ),
      selection,
      "page setup",
    );
    root.focus();
  }

  /** Format → Columns (SPEC-030 AC-1..AC-5): prefilled from the caret's
   * section columns and its content width. */
  function openColumnsDialog() {
    const model = doc.value;
    if (!model) {
      return;
    }
    const sectionIndex = Math.min(getCaretSectionIndex(), model.sections.length - 1);
    columnsDialogInitial.value = {
      columns: { ...sectionColumns(model, sectionIndex) },
      contentWidthCm: sectionContentWidthCm(model, sectionIndex),
      applyTo: "thisSection",
    };
    columnsDialogOpen.value = true;
  }

  /** The Columns dialog's OK (AC-6/AC-7): applies the layout to the caret's
   * section or wraps the selection in a continuous section, as one undo step. */
  function handleColumnsApply(
    payload = {
      columns: { count: 1, equalWidth: true, gapCm: 1.27, separator: false, widthsCm: undefined },
      applyTo: "thisSection",
    },
  ) {
    columnsDialogOpen.value = false;
    const root = contentRoot.value as HTMLElement | null;
    const model = doc.value;
    if (!root || !model || !editable.value) {
      return;
    }
    const selection = paragraphSelection(root);
    applyEdit(
      applyColumns(model, selection, payload.columns, asColumnsScope(payload.applyTo ?? "thisSection")),
      selection,
      "columns",
    );
    root.focus();
  }

  /** The "Show/hide ruler" toggle (AC-14): a view preference, not an
   * undo-able content edit — same direct-mutation + touchDocument pattern as
   * handleTitleUpdate, and triggerRef is required for the same reason (doc is
   * a shallowRef; mutating a field on its current value alone doesn't
   * re-render). */
  function toggleRuler() {
    const model = doc.value;
    if (!model) {
      return;
    }
    // Absent reads as shown (documentModel.ts's own convention), so the flip
    // is "was it explicitly false" rather than a plain `!model.showRuler`.
    model.showRuler = model.showRuler === false;
    triggerRef(doc);
    touchDocument(model.id);
  }

  /** Sets the document's measurement-unit preference from a ruler right-click
   * (SPEC-064 AC-10a). Writes the same `measurementUnit` field the Page setup
   * dialog writes (AC-10b); like the ruler toggle it is a persisted app
   * preference, not an undoable content edit, so it uses the same direct
   * mutation + triggerRef + touchDocument path. */
  function setMeasurementUnit(unit = "cm") {
    const model = doc.value;
    if (!model) {
      return;
    }
    model.measurementUnit = asMeasurementUnit(unit);
    triggerRef(doc);
    touchDocument(model.id);
  }

  // The ruler and the Page setup dialog follow the caret's section (AC-17..21)
  // so a per-section page setup is edited where the caret sits.
  const rulerSectionIndex = computed(() => {
    const model = doc.value;
    return model ? Math.min(getCaretSectionIndex(), model.sections.length - 1) : 0;
  });
  const rulerPageWidthCm = computed(() =>
    doc.value ? sectionPageSize(doc.value, rulerSectionIndex.value).widthCm : 21,
  );
  const rulerMarginLeftCm = computed(() =>
    doc.value ? sectionMargins(doc.value, rulerSectionIndex.value).leftCm : 2.54,
  );
  const rulerMarginRightCm = computed(() =>
    doc.value ? sectionMargins(doc.value, rulerSectionIndex.value).rightCm : 2.54,
  );
  // The caret section's column boundaries for the ruler (SPEC-030 AC-14): each
  // gutter's centre, measured from the content area's start; null when the
  // section has a single column (no boundary handles).
  const rulerColumns = computed(() => {
    const model = doc.value;
    if (!model) {
      return undefined;
    }
    const columns = sectionColumns(model, rulerSectionIndex.value);
    if (columns.count <= 1) {
      return undefined;
    }
    const widthsCm = resolveColumnWidthsCm(columns, sectionContentWidthCm(model, rulerSectionIndex.value));
    const boundaries = [];
    let offset = 0;
    for (let i = 0; i < widthsCm.length - 1; i += 1) {
      offset += widthsCm[i];
      boundaries.push({ index: i, centerCm: offset + columns.gapCm / 2 });
      offset += columns.gapCm;
    }
    return { boundaries };
  });

  return {
    pageSetupOpen,
    pageSetupInitial,
    columnsDialogOpen,
    columnsDialogInitial,
    openPageSetup,
    handlePageSetupApply,
    openColumnsDialog,
    handleColumnsApply,
    toggleRuler,
    setMeasurementUnit,
    rulerSectionIndex,
    rulerPageWidthCm,
    rulerMarginLeftCm,
    rulerMarginRightCm,
    rulerColumns,
  };
}
