// The sheet tab bar's commands (SPEC-121 "UI"): what each tab gesture, the
// tab context menu, the dialogs and the colour picker do — every command one
// undo step through useSpreadsheetUndo (record → lib/spreadsheet/sheets.ts
// command → commit, which runs the edit pipeline). Before recording, the
// current view state (each sheet's selection, the sheet on show) is written
// into the live model so the step's snapshot holds it; after a command the
// sheet its result names is shown.
//
// - Insert (the "+" button or the menu) adds `Sheet<N>` after the active
//   sheet. Rename edits in place: Enter commits, Escape cancels, blur commits
//   a valid name and cancels an invalid one, an invalid name keeps the input
//   open with its "Sheet names" message, and the unchanged name records
//   nothing. Delete asks first (SheetTabDeleteDialog) when the sheet has
//   content. Move left/right step past the adjacent visible tab; a drag drops
//   before a tab (or after the last). Each move is announced in the tab
//   bar's live region: `Sheet "<name>" moved to position <k>` (k over the
//   visible tabs). Hide, Unhide… (SheetUnhideDialog, "hidden" sheets only —
//   never "veryHidden") and Tab colour (the core TextColorPicker with the
//   recent colours of core/colors.ts, plus "No colour").
// - Opening a tab's menu (right-click, Shift+F10, ContextMenu) shows that
//   sheet first, as a spreadsheet does. Every item is present; one that
//   cannot apply is disabled with its reason.
// - SPEC-122 "Protection": a workbook whose structure is protected
//   (`structureProtected`) disables every sheet command with the reason
//   STRUCTURE_PROTECTED_MESSAGE — each menu item, and the gestures (the "+"
//   button, double-click rename, drag) refuse with it as a toast. The
//   protection XML itself is never changed.
// - SPEC-126: adding or duplicating a sheet in a csv/tsv workbook is allowed,
//   and the first time it happens in this view the toast
//   CSV_SECOND_SHEET_NOTICE says the file keeps only the first sheet.
import { computed, ref, type ComputedRef, type Ref, type ShallowRef } from "vue";
import { useToasts } from "../../composables/useToasts";
import { loadRecentColors, pushRecentColor } from "../../lib/core/colors";
import { DEFAULT_THEME, type ModelColor } from "../../lib/core/theme";
import { CSV_SECOND_SHEET_NOTICE } from "../../lib/spreadsheet/csv/messages";
import { isCsvFormat, type SpreadsheetModel } from "../../lib/spreadsheet/model";
import {
  addSheet,
  deleteSheet,
  duplicateSheet,
  isLastVisibleSheet,
  moveSheet,
  renameSheet,
  setSheetHidden,
  setTabColor,
  sheetHasContent,
  tabColorHex,
  validateSheetName,
  type SheetCommandResult,
} from "../../lib/spreadsheet/sheets";
import { movedAnnouncement, sheetTabMenuItems, STRUCTURE_PROTECTED_MESSAGE, visibleNeighbour, type SheetTabMenuItem } from "./sheetTabMenu";
import type { SpreadsheetUndo } from "./useSpreadsheetUndo";

export interface SheetCommandsDeps {
  book: {
    workbook: ShallowRef<SpreadsheetModel | null>;
    activeIndex: Ref<number>;
    selectSheet(index: number): void;
  };
  undo: SpreadsheetUndo;
  /** Writes the view state into the live model before a step is recorded. */
  syncViewState: (model: SpreadsheetModel) => void;
}

export function useSheetCommands(deps: SheetCommandsDeps) {
  const { addToast } = useToasts();
  const menu = ref<{ index: number; x: number; y: number } | null>(null);
  const renaming = ref<{ index: number; draft: string; error: string } | null>(null);
  const deleting = ref<{ index: number; name: string } | null>(null);
  const unhiding = ref(false);
  const coloring = ref<{ index: number; x: number; y: number } | null>(null);
  const announcement = ref("");
  const recentColors = ref(loadRecentColors());
  /** CSV_SECOND_SHEET_NOTICE has been shown in this view (SPEC-126). */
  let secondSheetNoticed = false;

  const model = (): SpreadsheetModel | null => deps.book.workbook.value;

  /** SPEC-126: once per view, a csv/tsv workbook that gained a sheet says
   * that only the first one is saved in the file. */
  function noticeSecondSheet(result: SheetCommandResult | undefined): void {
    const target = model();
    if (result && target && isCsvFormat(target.format) && !secondSheetNoticed) {
      secondSheetNoticed = true;
      addToast(CSV_SECOND_SHEET_NOTICE);
    }
  }

  /** Whether the workbook's structure is protected (SPEC-122) — and if so,
   * says why as a toast: the gesture is refused. */
  function refusedByProtection(): boolean {
    if (!model()?.structureProtected) {
      return false;
    }
    addToast(STRUCTURE_PROTECTED_MESSAGE);
    return true;
  }

  /** Runs one command as one undo step, then shows the sheet it names. */
  function run(label: string, command: (target: SpreadsheetModel) => SheetCommandResult | void): SheetCommandResult | undefined {
    const target = model();
    if (!target || refusedByProtection()) {
      return undefined;
    }
    deps.syncViewState(target);
    const active = target.sheets[deps.book.activeIndex.value];
    deps.undo.record(label, active ? { sheetId: active.id } : "workbook");
    let result: SheetCommandResult | void;
    try {
      result = command(target);
    } catch (error) {
      deps.undo.commit();
      addToast(error instanceof Error ? error.message : String(error));
      return undefined;
    }
    deps.undo.commit(result?.alsoRecalc ?? []);
    if (result) {
      deps.book.activeIndex.value = result.activeIndex;
    }
    return result ?? undefined;
  }

  const menuItems: ComputedRef<SheetTabMenuItem[]> = computed(() => {
    const target = model();
    return target && menu.value ? sheetTabMenuItems(target, menu.value.index) : [];
  });

  function openMenu(payload: { index: number; x: number; y: number }): void {
    deps.book.selectSheet(payload.index);
    menu.value = payload;
  }

  function closeMenu(): void {
    menu.value = null;
  }

  function insert(): void {
    noticeSecondSheet(run("Add sheet", (target) => addSheet(target, deps.book.activeIndex.value)));
  }

  // --- rename ---------------------------------------------------------------------------

  function startRename(index: number): void {
    const sheet = model()?.sheets[index];
    if (sheet && !refusedByProtection()) {
      renaming.value = { index, draft: sheet.name, error: "" };
    }
  }

  function renameInput(value: string): void {
    if (renaming.value) {
      renaming.value = { ...renaming.value, draft: value, error: "" };
    }
  }

  /** Enter: commits a valid name (the unchanged one records nothing); an
   * invalid one keeps the input open with its message. */
  function commitRename(): void {
    const target = model();
    const state = renaming.value;
    if (!target || !state) {
      return;
    }
    const name = state.draft.trim();
    if (name === target.sheets[state.index]?.name) {
      renaming.value = null;
      return;
    }
    const error = validateSheetName(target, name, state.index);
    if (error !== null) {
      renaming.value = { ...state, error };
      return;
    }
    renaming.value = null;
    run("Rename sheet", (each) => renameSheet(each, state.index, name));
  }

  function cancelRename(): void {
    renaming.value = null;
  }

  /** Blur commits a valid name and cancels an invalid one. */
  function blurRename(): void {
    const target = model();
    const state = renaming.value;
    if (!target || !state) {
      return;
    }
    if (validateSheetName(target, state.draft.trim(), state.index) === null) {
      commitRename();
    } else {
      cancelRename();
    }
  }

  // --- delete, hide, unhide ----------------------------------------------------------------

  function requestDelete(index: number): void {
    const target = model();
    if (!target || isLastVisibleSheet(target, index) || refusedByProtection()) {
      return;
    }
    if (sheetHasContent(target, index)) {
      deleting.value = { index, name: target.sheets[index]!.name };
    } else {
      run("Delete sheet", (each) => deleteSheet(each, index));
    }
  }

  function confirmDelete(): void {
    const state = deleting.value;
    deleting.value = null;
    if (state) {
      run("Delete sheet", (each) => deleteSheet(each, state.index));
    }
  }

  function cancelDelete(): void {
    deleting.value = null;
  }

  const hiddenSheets = computed(() =>
    (model()?.sheets ?? []).flatMap((sheet, index) => (sheet.visibility === "hidden" ? [{ index, name: sheet.name }] : [])),
  );

  function confirmUnhide(indices: number[]): void {
    unhiding.value = false;
    if (indices.length > 0) {
      run("Unhide sheet", (target) => setSheetHidden(target, indices, false));
    }
  }

  // --- move ---------------------------------------------------------------------------------

  /** The moved sheet's 1-based position over the visible tabs. */
  function announceMove(result: SheetCommandResult | undefined): void {
    const target = model();
    const sheet = result ? target?.sheets[result.activeIndex] : undefined;
    if (!target || !sheet) {
      return;
    }
    const position = target.sheets.slice(0, result!.activeIndex + 1).filter((each) => each.visibility === "visible").length;
    announcement.value = movedAnnouncement(sheet.name, position);
  }

  /** Move left/right: past the adjacent visible tab. */
  function step(index: number, direction: -1 | 1): void {
    const target = model();
    const neighbour = target ? visibleNeighbour(target, index, direction) : -1;
    if (neighbour >= 0) {
      announceMove(run("Move sheet", (each) => moveSheet(each, index, neighbour)));
    }
  }

  /** A drag: the sheet at `index` dropped before the sheet at `before`
   * (-1: after the last visible tab). */
  function dropTab(payload: { index: number; before: number }): void {
    const target = model();
    if (!target || !target.sheets[payload.index]) {
      return;
    }
    const remaining = target.sheets.filter((_, at) => at !== payload.index);
    let to: number;
    if (payload.before < 0) {
      const lastVisible = remaining.findLastIndex((sheet) => sheet.visibility === "visible");
      to = lastVisible + 1;
    } else {
      to = remaining.indexOf(target.sheets[payload.before]!);
    }
    if (to < 0 || to === payload.index) {
      return;
    }
    announceMove(run("Move sheet", (each) => moveSheet(each, payload.index, to)));
  }

  // --- tab colour ----------------------------------------------------------------------------

  function applyColor(color: ModelColor | null): void {
    const state = coloring.value;
    coloring.value = null;
    if (!state) {
      return;
    }
    if (typeof color === "string") {
      recentColors.value = pushRecentColor(color, recentColors.value);
    }
    run("Tab colour", (target) => setTabColor(target, state.index, color === null ? undefined : color));
  }

  function closeColor(): void {
    coloring.value = null;
  }

  const theme = computed(() => model()?.theme ?? DEFAULT_THEME);

  /** A tab colour as CSS, for the tab bar ("" for none). */
  function tabColorCss(color: ModelColor | undefined): string {
    const target = model();
    return target ? tabColorHex(target, color) : "";
  }

  /** The context menu's action. */
  function runMenu(action: string): void {
    const state = menu.value;
    menu.value = null;
    if (!state) {
      return;
    }
    switch (action) {
      case "insert":
        insert();
        break;
      case "rename":
        startRename(state.index);
        break;
      case "duplicate":
        noticeSecondSheet(run("Duplicate sheet", (target) => duplicateSheet(target, state.index)));
        break;
      case "delete":
        requestDelete(state.index);
        break;
      case "move-left":
        step(state.index, -1);
        break;
      case "move-right":
        step(state.index, 1);
        break;
      case "hide":
        run("Hide sheet", (target) => setSheetHidden(target, [state.index], true));
        break;
      case "unhide":
        unhiding.value = !refusedByProtection();
        break;
      case "tab-color":
        if (!refusedByProtection()) {
          coloring.value = { index: state.index, x: state.x, y: state.y };
        }
        break;
      default:
        break;
    }
  }

  /** Why the sheet commands are disabled (SPEC-122); "" when they apply. */
  const structureReason = computed(() => (model()?.structureProtected ? STRUCTURE_PROTECTED_MESSAGE : ""));

  return {
    menu,
    menuItems,
    structureReason,
    openMenu,
    closeMenu,
    runMenu,
    insert,
    renaming,
    startRename,
    renameInput,
    commitRename,
    cancelRename,
    blurRename,
    deleting,
    requestDelete,
    confirmDelete,
    cancelDelete,
    unhiding,
    hiddenSheets,
    confirmUnhide,
    cancelUnhide: () => {
      unhiding.value = false;
    },
    dropTab,
    announcement,
    coloring,
    applyColor,
    closeColor,
    recentColors,
    theme,
    tabColorCss,
  };
}

export type SheetCommands = ReturnType<typeof useSheetCommands>;
