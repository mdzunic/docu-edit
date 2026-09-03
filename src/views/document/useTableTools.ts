// The Table menu's feature logic (SPEC-076): the header-row toggle, the
// style gallery and its look flags, row sorting, and both text↔table
// conversions.
//
// Every rule this drives is pure and lives elsewhere — tableStyles.ts,
// tableSort.ts, tableConvert.ts for the arithmetic (AC-78), and
// commands.tableTools.ts for the model surgery. What is here is the UI
// layer's own share: which commands exist, whether each applies right now
// and why not, which dialog is open, and the one-undo-entry applyEdit call
// each action makes (AC-79).
import { computed, ref } from "vue";
import { caretTableLocation, setRowIsHeader, tableAt, type TableLocation, type TableMenuState } from "../../lib/tables";
import {
  applyTableStyleLook,
  convertibleParagraphRange,
  convertTableToText,
  convertTextToTable,
  setTableLook,
  sortTable,
} from "../../lib/commands.tableTools";
import { BUILT_IN_TABLE_STYLES, leadingHeaderRowCount } from "../../lib/tableStyles";
import { noSortKeys, sortRangeHasMergedRows, sortRangeRowCount, type SortKey } from "../../lib/tableSort";
import { separatorCharacter, type ConvertSeparatorId } from "../../lib/tableConvert";
import { caretFromRepeatedHeaderPoint } from "../../lib/domCaret";
import { DEFAULT_LANGUAGE } from "../../lib/core/language";
import { caretSelection, type ModelSelection } from "../../lib/commands";
import type { Ref } from "vue";
import type { CaretPosition } from "../../lib/textEditing";
import type { Table, TableLook } from "../../lib/documentModel";
import type { TableCellTarget } from "../../lib/domCaret";
import type { EditorContext } from "./context";

/** Why a Table command does not apply right now — shown in the menu item
 * itself, because SPEC-049 disables commands rather than hiding them
 * (AC-18). */
export const NOT_IN_TABLE_REASON = "the caret is not in a table";
export const NO_PARAGRAPHS_REASON = "select one or more paragraphs";

/** Which Table dialog is open; "" is none. Kept as one value rather than
 * four booleans so the host component switches on it and DocumentView's
 * template hookup stays a single element (AC-79). */
export type TableDialog = "" | "gallery" | "sort" | "text-to-table" | "table-to-text";

export interface TableToolsDeps extends EditorContext {
  /** DocumentView's caret mirror — the reactive source the menu's enabled
   * state and check marks read (the raw getCaret() is not reactive). */
  caretPosition: Ref<CaretPosition | null>;
  /** DocumentView's paragraphSelection: always resolves to a selection. */
  paragraphSelection: (root?: HTMLElement | null) => ModelSelection;
  /** The cell a context menu was opened on (SPEC-022's shared state) — the
   * row "Repeat header row" acts on in the context menu (AC-19). */
  getContextMenuTableCell: () => TableCellTarget | null;
  closeContextMenu: () => void;
}

export function useTableTools(deps: TableToolsDeps) {
  const { doc, contentRoot, editable, applyEdit, applyCaretMove, addToast } = deps;

  const dialog = ref("" as TableDialog);

  // --- where the caret is ---------------------------------------------------

  const caretLocation = computed<TableLocation | null>(() => {
    const model = doc.value;
    return model ? caretTableLocation(model, deps.caretPosition.value) : null;
  });

  /** The document's style catalogue — the gallery previews through it and
   * marks the applied entry selected (AC-32, AC-34). */
  const docStyles = computed(() => doc.value?.styles ?? []);

  const caretTable = computed<Table | null>(() => {
    const model = doc.value;
    const location = caretLocation.value;
    return model && location ? tableAt(model, location.section, location.block) : null;
  });

  /** The paragraph range "Convert text to table" would consume, or null
   * (AC-52). Read from the live selection at menu-open time. */
  function convertibleRange() {
    const model = doc.value;
    if (!model) {
      return null;
    }
    return convertibleParagraphRange(model, deps.paragraphSelection(contentRoot.value as HTMLElement | null));
  }

  /** Whether the caret row is a header row — the "Repeat header row" check
   * item's state (AC-17). */
  const caretRowIsHeader = computed(() => {
    const location = caretLocation.value;
    return !!location && !!caretTable.value?.rows[location.row]?.isHeader;
  });

  /** What the menu bar needs to draw the Table menu: one object so
   * DocumentView passes one prop (AC-79). */
  const menuState = computed<TableMenuState>(() => ({
    inTable: caretTable.value !== null,
    rowIsHeader: caretRowIsHeader.value,
    reason: NOT_IN_TABLE_REASON,
    // "Convert text to table" is the one command whose condition is about
    // the selection, not the caret's table (AC-52).
    canConvertText: caretTable.value === null && convertibleRange() !== null,
    convertTextReason: caretTable.value === null ? NO_PARAGRAPHS_REASON : "the selection is inside a table",
  }));

  // --- header rows (AC-17, AC-19) -------------------------------------------

  function setHeaderRow(location: TableLocation, isHeader: boolean) {
    const model = doc.value;
    if (!model || !editable.value) {
      return;
    }
    const before = caretSelection({
      section: location.section,
      block: location.block,
      offset: 0,
      cell: { row: location.row, col: location.col, block: 0 },
    });
    applyEdit(
      setRowIsHeader(model, before, location, location.row, isHeader),
      before,
      isHeader ? "repeat header row" : "stop repeating header row",
    );
  }

  /** Table → Repeat header row: toggles the CARET row's isHeader (AC-17). */
  function toggleRepeatHeaderRow() {
    const location = caretLocation.value;
    if (!location) {
      addToast("Put the caret in a table first");
      return;
    }
    setHeaderRow(location, !caretRowIsHeader.value);
  }

  /** The same check item in the table section of the context menu, acting on
   * the RIGHT-CLICKED row rather than the caret's (AC-19). */
  function toggleContextRepeatHeaderRow() {
    const target = deps.getContextMenuTableCell();
    const model = doc.value;
    deps.closeContextMenu();
    if (!target || !model) {
      return;
    }
    const table = tableAt(model, target.section, target.block);
    setHeaderRow(target, !table?.rows[target.row]?.isHeader);
  }

  /** Whether the right-clicked row is a header row — the context menu's own
   * check state (AC-19). */
  function contextRowIsHeader(): boolean {
    const target = deps.getContextMenuTableCell();
    const model = doc.value;
    if (!target || !model) {
      return false;
    }
    return !!tableAt(model, target.section, target.block)?.rows[target.row]?.isHeader;
  }

  // --- the style gallery (AC-31..AC-36) -------------------------------------

  function openStyleGallery() {
    if (!caretTable.value) {
      addToast("Put the caret in a table first");
      return;
    }
    dialog.value = "gallery";
  }

  /** Applies a gallery tile in one undo entry (AC-33); `styleId` null is the
   * gallery's "None" entry (AC-34). */
  function applyGalleryStyle(styleId: string | null) {
    const model = doc.value;
    const location = caretLocation.value;
    if (!model || !location || !editable.value) {
      return;
    }
    const entry = BUILT_IN_TABLE_STYLES.find((style) => style.id === styleId);
    const before = caretSelection({ section: location.section, block: location.block, offset: 0 });
    applyEdit(
      applyTableStyleLook(model, before, location, styleId, entry ? { ...entry.look } : undefined),
      before,
      "table style",
    );
  }

  /** One look checkbox: takes effect immediately, one undo entry (AC-35). */
  function toggleLookFlag(key: keyof TableLook) {
    const model = doc.value;
    const location = caretLocation.value;
    const table = caretTable.value;
    if (!model || !location || !table || !editable.value) {
      return;
    }
    const look: TableLook = { ...table.look };
    if (look[key]) {
      delete look[key];
    } else {
      look[key] = true;
    }
    const before = caretSelection({ section: location.section, block: location.block, offset: 0 });
    applyEdit(setTableLook(model, before, location, look), before, "table banding");
  }

  // --- sorting (AC-38..AC-41, AC-49..AC-51) ---------------------------------

  const sortHasHeader = ref(true);
  const sortKeys = ref(noSortKeys());

  function openSortDialog() {
    const table = caretTable.value;
    const location = caretLocation.value;
    if (!table || !location) {
      addToast("Put the caret in a table first");
      return;
    }
    // AC-39: key one defaults to the caret's column, ascending, text; keys
    // two and three default to "(none)", which the dialog shows as an
    // unset column picker.
    sortKeys.value = [{ columnIndex: location.col, direction: "asc", type: "text" }];
    // AC-40: the header toggle defaults to whether the first row is one.
    sortHasHeader.value = leadingHeaderRowCount(table) > 0;
    dialog.value = "sort";
  }

  /** Why the sort dialog's OK is disabled, or "" when it applies (AC-49,
   * AC-50). Recomputed as the header toggle moves the sort range. */
  function sortBlockedReason(hasHeader: boolean): string {
    const table = caretTable.value;
    if (!table) {
      return NOT_IN_TABLE_REASON;
    }
    if (sortRangeRowCount(table, hasHeader) < 2) {
      return "there is nothing to sort";
    }
    if (sortRangeHasMergedRows(table, hasHeader)) {
      return "sorting is not available for tables with merged rows";
    }
    return "";
  }

  function applySort(keys: SortKey[], hasHeader: boolean) {
    const model = doc.value;
    const location = caretLocation.value;
    dialog.value = "";
    if (!model || !location || !editable.value || sortBlockedReason(hasHeader)) {
      return;
    }
    const before = caretSelection({ section: location.section, block: location.block, offset: 0 });
    applyEdit(
      sortTable(model, before, location, keys, hasHeader, model.defaultStyles.language ?? DEFAULT_LANGUAGE),
      before,
      "sort table",
    );
  }

  // --- the conversions (AC-52..AC-64) ---------------------------------------

  function openTextToTable() {
    if (!convertibleRange()) {
      addToast("Select one or more paragraphs first");
      return;
    }
    dialog.value = "text-to-table";
  }

  function openTableToText() {
    if (!caretTable.value) {
      addToast("Put the caret in a table first");
      return;
    }
    dialog.value = "table-to-text";
  }

  /** The paragraphs the text→table preview counts rows and columns over
   * (AC-54) — read live so the dialog re-counts as the separator changes. */
  const convertSourceParagraphs = computed(() => {
    const model = doc.value;
    const range = model ? convertibleParagraphRange(model, deps.paragraphSelection(contentRoot.value as HTMLElement | null)) : null;
    if (!model || !range) {
      return [];
    }
    return model.sections[range.section].blocks.slice(range.first, range.last + 1);
  });

  function applyTextToTable(id: ConvertSeparatorId, custom: string) {
    const model = doc.value;
    const selection = deps.paragraphSelection(contentRoot.value as HTMLElement | null);
    dialog.value = "";
    if (!model || !editable.value) {
      return;
    }
    applyEdit(
      convertTextToTable(model, selection, separatorCharacter(id, custom)),
      selection,
      "convert text to table",
    );
  }

  function applyTableToText(id: ConvertSeparatorId, custom: string) {
    const model = doc.value;
    const location = caretLocation.value;
    dialog.value = "";
    if (!model || !location || !editable.value) {
      return;
    }
    const before = caretSelection({ section: location.section, block: location.block, offset: 0 });
    applyEdit(
      convertTableToText(model, before, location, separatorCharacter(id, custom)),
      before,
      "convert table to text",
    );
  }

  // --- the repeated header run's clicks (AC-8) ------------------------------

  /** A mousedown on a repeated header row places the caret in the REAL
   * header cell it copies, at the offset nearest the click. Returns true
   * when it consumed the event, so the canvas hub stops there rather than
   * letting the browser drop the caret at whatever editable position lies
   * next to an uneditable row. */
  function redirectRepeatedHeaderClick(event: MouseEvent): boolean {
    const root = contentRoot.value as HTMLElement | null;
    if (!root || !(event.target instanceof Node)) {
      return false;
    }
    // The repeat's own DOM carries the real cell's address, so the nearest
    // text position inside the copy resolves straight to the original.
    const point = caretPointFromClick(root, event);
    const caret = point ? caretFromRepeatedHeaderPoint(root, point.node, point.offset) : null;
    if (!caret) {
      return false;
    }
    event.preventDefault();
    applyCaretMove(caret);
    return true;
  }

  /** The [id, handler] pairs DocumentView spreads into its editorCommands
   * registry, so every Table command is dispatchable from the palette and
   * the menu alike (AC-80). Declared in the shape commandCatalogue.spec.ts's
   * completeness guard reads, like every other composable that contributes
   * command ids. */
  const commands: Array<[string, () => void]> = [
    ["table.repeat-header-row", toggleRepeatHeaderRow],
    ["table.style-gallery", openStyleGallery],
    ["table.sort", openSortDialog],
    ["table.convert-text-to-table", openTextToTable],
    ["table.convert-table-to-text", openTableToText],
  ];

  return {
    dialog,
    menuState,
    caretTable,
    docStyles,
    caretRowIsHeader,
    contextRowIsHeader,
    convertSourceParagraphs,
    sortKeys,
    sortHasHeader,
    setSortHasHeader: (value: boolean) => {
      sortHasHeader.value = value;
    },
    sortBlockedReason,
    toggleRepeatHeaderRow,
    toggleContextRepeatHeaderRow,
    openStyleGallery,
    applyGalleryStyle,
    toggleLookFlag,
    openSortDialog,
    applySort,
    openTextToTable,
    openTableToText,
    applyTextToTable,
    applyTableToText,
    redirectRepeatedHeaderClick,
    closeDialog: () => {
      dialog.value = "";
    },
    commands,
  };
}

/** The DOM text point a click landed on, using whichever of the two browser
 * APIs exists (Chrome's caretRangeFromPoint, Firefox's caretPositionFromPoint).
 * Returns null where neither is available — jsdom, and any environment that
 * cannot hit-test. */
function caretPointFromClick(
  root: Element,
  event: MouseEvent,
): { node: Node; offset: number } | null {
  const view = root.ownerDocument.defaultView as (Window & {
    // Neither API is in the DOM lib's Document type across all targets.
    document: Document & {
      caretRangeFromPoint?: (x: number, y: number) => Range | null;
      caretPositionFromPoint?: (x: number, y: number) => { offsetNode: Node; offset: number } | null;
    };
  }) | null;
  const document_ = view?.document;
  if (!document_) {
    return null;
  }
  const range = document_.caretRangeFromPoint?.(event.clientX, event.clientY);
  if (range) {
    return { node: range.startContainer, offset: range.startOffset };
  }
  const position = document_.caretPositionFromPoint?.(event.clientX, event.clientY);
  return position ? { node: position.offsetNode, offset: position.offset } : null;
}
