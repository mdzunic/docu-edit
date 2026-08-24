// The vertical ruler down the left edge of the page (SPEC-064): its
// caret-section geometry inputs, the caret table's row boundaries (measured
// live from the DOM), the "which band is being edited" flag, its print-layout
// visibility, and the drag handler that turns an emitted proposal into an
// undoable model edit. Kept out of the DocumentView backbone — like
// usePageSetupDialogs holds the horizontal ruler's geometry — so the view
// gains only a `use*({...})` wiring call and the template hookup, staying
// within its line budget (DocumentView.size.spec.ts).
//
// The ruler is presentational: it emits `update`; nothing here reaches into it.
// Margin/band edits go straight to the document model (they are page-level, not
// text-surface edits) via applyEdit, mirroring the horizontal ruler's
// per-drag-step undo granularity (AC-2a). A table-row drag reuses SPEC-023's
// setRowMinHeight under the same "resize row" label (AC-5a).
import { computed, ref, watch } from "vue";
import { caretSelection, setHeaderFooterDistance, setPageMargins } from "../../lib/commands";
import { CM_PER_PX } from "../../lib/core/units";
import { tableRowElement } from "../../lib/core/tables";
import {
  DEFAULT_HEADER_FOOTER_DISTANCE_CM,
  sectionMargins,
  sectionPageSize,
} from "../../lib/documentModel";
import { setRowMinHeight, tableAt } from "../../lib/tables";
import { findTableElement } from "../../lib/tableResize";
import type { VerticalRowBoundary } from "../../lib/verticalRuler";
import type { Ref } from "vue";
import type { ModelSelection } from "../../lib/commands";
import type { ActiveHeaderFooterTarget } from "../../lib/documentModel";
import type { CaretPosition } from "../../lib/textEditing";
import type { EditorContext } from "./context";
import type { ViewMode } from "./useViewModes";

export interface VerticalRulerDeps extends EditorContext {
  /** DocumentView's paragraphSelection always resolves to a body selection. */
  paragraphSelection: (root?: HTMLElement | null) => ModelSelection;
  /** Live read of the caret-section mirror, as usePageSetupDialogs receives. */
  getCaretSectionIndex: () => number;
  /** The header/footer editing target, or null in the body (SPEC-029) — its
   * `part` names the band the ruler must mark as active (AC-6). */
  activeHeaderFooter: Ref<ActiveHeaderFooterTarget | null>;
  /** The current layout: the vertical ruler shows only in print layout
   * (AC-8). */
  viewMode: Readonly<Ref<ViewMode>>;
  /** The reactive caret mirror — its change re-measures the caret table's rows
   * (AC-5), since a caret move can enter or leave a table without an edit. */
  caretPosition: Ref<CaretPosition | null>;
}

export function useVerticalRuler(deps: VerticalRulerDeps) {
  const {
    doc,
    contentRoot,
    editable,
    applyEdit,
    currentSelection,
    clientToContentPx,
    zoom,
    getCaret,
    onScroll,
    paragraphSelection,
    getCaretSectionIndex,
    activeHeaderFooter,
    viewMode,
    caretPosition,
  } = deps;

  const sectionIndex = computed(() => {
    const model = doc.value;
    return model ? Math.min(getCaretSectionIndex(), model.sections.length - 1) : 0;
  });

  const pageWidthCm = computed(() => (doc.value ? sectionPageSize(doc.value, sectionIndex.value).widthCm : 21));
  const pageHeightCm = computed(() => (doc.value ? sectionPageSize(doc.value, sectionIndex.value).heightCm : 29.7));
  const marginTopCm = computed(() => (doc.value ? sectionMargins(doc.value, sectionIndex.value).topCm : 2.54));
  const marginBottomCm = computed(() => (doc.value ? sectionMargins(doc.value, sectionIndex.value).bottomCm : 2.54));
  const headerDistanceCm = computed(() =>
    doc.value ? (doc.value.sections[sectionIndex.value]?.headerDistanceCm ?? DEFAULT_HEADER_FOOTER_DISTANCE_CM) : DEFAULT_HEADER_FOOTER_DISTANCE_CM,
  );
  const footerDistanceCm = computed(() =>
    doc.value ? (doc.value.sections[sectionIndex.value]?.footerDistanceCm ?? DEFAULT_HEADER_FOOTER_DISTANCE_CM) : DEFAULT_HEADER_FOOTER_DISTANCE_CM,
  );

  // "body" | "header" | "footer" — the band being edited (AC-6).
  const activeBand = computed(() => activeHeaderFooter.value?.part ?? "body");

  const unit = computed(() => doc.value?.measurementUnit ?? "cm");

  // --- caret table row boundaries (AC-5) ------------------------------------
  const tableRows = ref<VerticalRowBoundary[]>([]);

  /** Re-measures the caret table's row boundaries from the DOM as distances
   * from the page's top edge (through SPEC-063's client→content map so they are
   * zoom-correct). Empty when the caret is not inside a table. */
  function measureRows() {
    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    const caret = getCaret();
    if (!model || !root || !caret || !caret.cell) {
      tableRows.value = [];
      return;
    }
    const table = tableAt(model, caret.section, caret.block);
    const tableEl = findTableElement(root, caret.section, caret.block);
    if (!table || !tableEl) {
      tableRows.value = [];
      return;
    }
    const boundaries: VerticalRowBoundary[] = [];
    for (let i = 0; i < table.rows.length; i += 1) {
      const trEl = tableRowElement(tableEl, i);
      if (!trEl) {
        continue;
      }
      const rect = trEl.getBoundingClientRect();
      boundaries.push({
        row: i,
        topCm: clientToContentPx(0, rect.top).y * CM_PER_PX,
        bottomCm: clientToContentPx(0, rect.bottom).y * CM_PER_PX,
      });
    }
    tableRows.value = boundaries;
  }

  // A caret move (no edit) or any edit can change which rows are shown; both
  // must re-measure after the DOM settles. Scroll shifts every rect too.
  watch([caretPosition, doc], () => measureRows(), { flush: "post" });
  onScroll(measureRows);

  // Its own per-document flag (View → Vertical ruler; View → Rulers sets both
  // flags at once), and print layout only (AC-8).
  const showVerticalRuler = computed(() => {
    const model = doc.value;
    return !!model && model.showVerticalRuler !== false && viewMode.value === "print";
  });

  const verticalRulerProps = computed(() => ({
    pageHeightCm: pageHeightCm.value,
    pageWidthCm: pageWidthCm.value,
    marginTopCm: marginTopCm.value,
    marginBottomCm: marginBottomCm.value,
    headerDistanceCm: headerDistanceCm.value,
    footerDistanceCm: footerDistanceCm.value,
    tableRows: tableRows.value,
    activeBand: activeBand.value,
    unit: unit.value,
    zoom: zoom.value,
    disabled: !editable.value,
    clientToContentPx,
  }));

  /** A vertical-ruler drag (AC-2a, AC-3a/3b, AC-5a): the emitted proposal
   * becomes an undoable model edit. Margins/band distances are page-level, so
   * they apply straight to the document; a row height reuses SPEC-023's
   * setRowMinHeight with its own "resize row" label. */
  function handleVerticalRulerUpdate(update = { kind: "", valueCm: 0, row: 0, minHeightCm: 0 }) {
    const model = doc.value;
    const root = contentRoot.value as HTMLElement | null;
    if (!model || !root || !editable.value) {
      return;
    }
    const selection = paragraphSelection(root);
    // Every branch passes skipCaretScroll: a drag fires one edit per pointermove
    // and the marker being dragged may be a full page below the caret, so the
    // user has scrolled it into view. keepCaretVisible would snap the viewport
    // back to the caret mid-drag and desync the pointer→content mapping
    // (SPEC-064 AC-2a/3b) — see ApplyEditOptions.skipCaretScroll.
    const dragOptions = { skipCaretScroll: true };
    if (update.kind === "marginTopCm") {
      applyEdit(setPageMargins(model, selection, { topCm: update.valueCm }), selection, "margins", dragOptions);
    } else if (update.kind === "marginBottomCm") {
      applyEdit(setPageMargins(model, selection, { bottomCm: update.valueCm }), selection, "margins", dragOptions);
    } else if (update.kind === "headerDistanceCm") {
      applyEdit(setHeaderFooterDistance(model, selection, { headerDistanceCm: update.valueCm }), selection, "header distance", dragOptions);
    } else if (update.kind === "footerDistanceCm") {
      applyEdit(setHeaderFooterDistance(model, selection, { footerDistanceCm: update.valueCm }), selection, "footer distance", dragOptions);
    } else if (update.kind === "table-row") {
      const caret = getCaret();
      if (!caret || !caret.cell) {
        return;
      }
      const location = { section: caret.section, block: caret.block, row: update.row, col: 0 };
      const rowSelection =
        currentSelection(root) ??
        caretSelection({ section: caret.section, block: caret.block, offset: 0, cell: { row: update.row, col: 0, block: 0 } });
      applyEdit(setRowMinHeight(model, rowSelection, location, update.row, update.minHeightCm), rowSelection, "resize row", dragOptions);
    }
  }

  return {
    showVerticalRuler,
    verticalRulerProps,
    handleVerticalRulerUpdate,
  };
}
