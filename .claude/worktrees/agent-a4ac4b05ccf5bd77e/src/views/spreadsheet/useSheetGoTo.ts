// Go To and Go To Special in the workbook view (SPEC-134 "View wiring"):
// GoToDialog's and GoToSpecialDialog's state, the jump every Go To and Name
// Box entry makes, and the keys. lib/spreadsheet/find.ts decides what a
// reference means (resolveGoToTarget) and what Go To Special selects
// (selectSpecial); sheetJump.ts selects it.
//
// - Keys (window, capture): Control+G on every platform and F5 — `sheet.goTo`
//   in SHEET_SHORTCUTS — open the dialog with preventDefault, so F5 never
//   reloads the page while a workbook view is mounted and no other dialog is
//   open (inside Go To's own dialogs they are consumed); while the cell
//   editor is open they are consumed and do nothing (Excel disables Go To in
//   edit mode).
// - The dialog: the reference field (pre-filled with the last origin, so F5
//   then Enter goes back, as in Excel), the list — the names and tables the
//   Name Box lists (nameBoxOptions: visible from the active sheet, hidden and
//   reserved names never), then the last GO_TO_RECENT_LIMIT origins, most
//   recent first — "Special…" and OK. Choosing an item fills the field;
//   activating it (Enter, double-click) goes. An error shows in the dialog
//   and leaves the selection unchanged.
// - A jump (Go To, and the Name Box's through useNames' setJump): the
//   previous selection is pushed onto the recent list first, as `Sheet!A1:B2`
//   text (goToReferenceText), then the target's ranges are selected with the
//   active cell at the first range's top-left, the sheet switched when needed
//   (view state, never dirty) and the target revealed.
// - Go To Special: Blanks, Constants, Formulas, Errors, Last cell, Visible
//   cells only, with the four type filters (all on by default, used by
//   Constants and Formulas); OK selects selectSpecial's rectangles, the
//   active cell at the first one's top-left, or shows why nothing changed.
//   The kind and the filters persist for the view session.
import { computed, onScopeDispose, ref, type Ref, type ShallowRef } from "vue";
import { isMacPlatform } from "../../lib/core/platform";
import { findShortcut } from "../../lib/shortcuts";
import type { GridRange } from "../../lib/spreadsheet/address";
import { goToReferenceText, GO_TO_INVALID_MESSAGE, resolveGoToTarget, selectSpecial, type SpecialKind, type SpecialTypes } from "../../lib/spreadsheet/find";
import type { Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { nameBoxOptions } from "./nameBox";
import { createSheetJump, type JumpOptions, type SheetJump } from "./sheetJump";
import { SHEET_SHORTCUTS } from "./sheetShortcuts";
import type { CellEditing } from "./useCellEditing";
import type { Names } from "./useNames";
import type { SheetSelectionController } from "./useSheetSelection";
import type { SheetViewState } from "./useSheetViewState";

/** How many Go To origins the list remembers. */
export const GO_TO_RECENT_LIMIT = 10;

export interface SheetGoToDeps {
  book: {
    workbook: ShallowRef<SpreadsheetModel | null>;
    activeIndex: Ref<number>;
    activeSheet: Readonly<Ref<Sheet | null>>;
    selectSheet(index: number): void;
  };
  viewState: Pick<SheetViewState, "update">;
  selection: Pick<SheetSelectionController, "selection">;
  editing: { cell: Pick<CellEditing, "editing" | "focusCell" | "revealTick"> };
  /** SPEC-133's names: the Name Box's jumps go through this one. */
  names?: Pick<Names, "setJump"> | null;
  /** Test seam: the platform. */
  isMac?: boolean;
}

/** One entry of the dialog's list: a name, a table or a recent origin. */
export interface GoToItem {
  key: string;
  kind: "name" | "table" | "recent";
  /** What the reference field takes. */
  text: string;
}

const ALL_TYPES: SpecialTypes = { numbers: true, text: true, logicals: true, errors: true };
/** The kinds whose cells the type filters narrow. */
export const TYPED_SPECIAL_KINDS: readonly SpecialKind[] = ["constants", "formulas"];

export function useSheetGoTo(deps: SheetGoToDeps) {
  const model = (): SpreadsheetModel | null => deps.book.workbook.value;
  const base = createSheetJump({ book: deps.book, viewState: deps.viewState, cell: deps.editing.cell });
  /** The origins, most recent first. */
  const recent = ref<string[]>([]);

  /** Pushes where the selection is now onto the recent list. */
  function remember(): void {
    const sheet = deps.book.activeSheet.value;
    if (sheet?.kind !== "grid") {
      return;
    }
    const origin = goToReferenceText(sheet.name, deps.selection.selection.value.ranges);
    recent.value = [origin, ...recent.value.filter((text) => text !== origin)].slice(0, GO_TO_RECENT_LIMIT);
  }

  /** Every Go To and Name Box jump: the origin remembered first. */
  const jump: SheetJump = async (sheetId: string, ranges: GridRange[], options?: JumpOptions) => {
    remember();
    await base(sheetId, ranges, options);
  };
  deps.names?.setJump(jump);

  // --- the Go To dialog -------------------------------------------------------------------------
  const dialogOpen = ref(false);
  const reference = ref("");
  const error = ref("");
  const chosen = ref<string | null>(null);

  const items = computed((): GoToItem[] => {
    const m = model();
    const sheet = deps.book.activeSheet.value;
    const named: GoToItem[] = m && sheet ? nameBoxOptions(m, sheet.id).map((option) => ({ key: option.key, kind: option.kind, text: option.name })) : [];
    return [...named, ...recent.value.map((text, index): GoToItem => ({ key: `recent:${index}`, kind: "recent", text }))];
  });

  function openDialog(): void {
    if (!model() || deps.editing.cell.editing.value) {
      return;
    }
    specialOpen.value = false;
    reference.value = recent.value[0] ?? "";
    error.value = "";
    chosen.value = null;
    dialogOpen.value = true;
  }

  function cancel(): void {
    dialogOpen.value = false;
    deps.editing.cell.focusCell();
  }

  function setReference(text: string): void {
    reference.value = text;
    error.value = "";
    chosen.value = null;
  }

  /** A list item chosen (a click): the field takes its text. */
  function choose(key: string): void {
    const item = items.value.find((each) => each.key === key);
    if (item) {
      setReference(item.text);
      chosen.value = key;
    }
  }

  /** OK (or Enter): the reference, resolved — a jump, or the error shown. */
  async function confirm(): Promise<boolean> {
    const m = model();
    const sheet = deps.book.activeSheet.value;
    if (!m || !sheet) {
      return false;
    }
    const target = reference.value.trim() === "" ? { error: GO_TO_INVALID_MESSAGE } : resolveGoToTarget(m, reference.value, sheet.id, deps.selection.selection.value.active);
    if ("error" in target) {
      error.value = target.error;
      return false;
    }
    dialogOpen.value = false;
    await jump(target.sheetId, target.ranges);
    return true;
  }

  /** A list item activated (Enter, a double-click): chosen, then gone to. */
  async function activate(key: string): Promise<boolean> {
    choose(key);
    return chosen.value === key ? confirm() : false;
  }

  // --- Go To Special ---------------------------------------------------------------------------
  const specialOpen = ref(false);
  const specialKind = ref<SpecialKind>("blanks");
  const specialTypes = ref<SpecialTypes>({ ...ALL_TYPES });
  const specialMessage = ref("");
  const typesEnabled = computed(() => TYPED_SPECIAL_KINDS.includes(specialKind.value));

  /** "Special…": Go To gives way to Go To Special. */
  function openSpecial(): void {
    dialogOpen.value = false;
    specialMessage.value = "";
    specialOpen.value = true;
  }

  function cancelSpecial(): void {
    specialOpen.value = false;
    deps.editing.cell.focusCell();
  }

  function setKind(kind: SpecialKind): void {
    specialKind.value = kind;
    specialMessage.value = "";
  }

  function setType(type: keyof SpecialTypes, on: boolean): void {
    specialTypes.value = { ...specialTypes.value, [type]: on };
    specialMessage.value = "";
  }

  /** OK: selectSpecial's rectangles selected, or the message shown. */
  async function confirmSpecial(): Promise<boolean> {
    const m = model();
    const sheet = deps.book.activeSheet.value;
    if (!m || sheet?.kind !== "grid") {
      return false;
    }
    const found = selectSpecial(m, sheet, deps.selection.selection.value, specialKind.value, specialTypes.value);
    if ("error" in found) {
      specialMessage.value = found.error;
      return false;
    }
    specialOpen.value = false;
    await base(sheet.id, found);
    return true;
  }

  // --- keys ------------------------------------------------------------------------------------
  function onKeydown(event: KeyboardEvent): void {
    if (event.defaultPrevented || findShortcut(event, deps.isMac ?? isMacPlatform(), SHEET_SHORTCUTS)?.id !== "sheet.goTo") {
      return;
    }
    const ours = dialogOpen.value || specialOpen.value;
    if (!ours && typeof document !== "undefined" && document.querySelector('[role="dialog"]') !== null) {
      return;
    }
    event.preventDefault();
    if (!ours && !deps.editing.cell.editing.value) {
      openDialog();
    }
  }
  if (typeof window !== "undefined") {
    window.addEventListener("keydown", onKeydown, true);
  }
  onScopeDispose(() => {
    if (typeof window !== "undefined") {
      window.removeEventListener("keydown", onKeydown, true);
    }
  });

  return {
    recent,
    jump,
    dialogOpen,
    reference,
    error,
    chosen,
    items,
    openDialog,
    cancel,
    setReference,
    choose,
    confirm,
    activate,
    specialOpen,
    specialKind,
    specialTypes,
    specialMessage,
    typesEnabled,
    openSpecial,
    cancelSpecial,
    setKind,
    setType,
    confirmSpecial,
    onKeydown,
  };
}

export type SheetGoTo = ReturnType<typeof useSheetGoTo>;

/** Narrows a component's runtime-typed `goTo` prop (the .vue parser bars casts). */
export function asSheetGoTo(value: unknown): SheetGoTo {
  return value as SheetGoTo;
}
