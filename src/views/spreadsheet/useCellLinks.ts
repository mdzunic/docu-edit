// SPEC-135's hyperlinks on cells in the workbook view (cellLinks.ts holds the
// framework-free parts; annotations.ts the commands and the open rule):
// - The link dialog (SheetLinkDialog, `dialog`): Ctrl/Cmd+K with the grid
//   focused and no cell being edited (`sheet.insertLink` in SHEET_SHORTCUTS),
//   Insert ▸ "Link…" and the cell context menu's "Link…" open it on the
//   active cell — its display text, and its stored link's target and tooltip
//   when it has one ("Edit link"). Save is ONE undo step, "Insert link" or
//   "Edit link": the text, when it differs from the cell's display, is
//   entered into the active cell through SPEC-122's cell-entry path
//   (entryFromText, setCellEntry, the row refit — a refused or unparseable
//   entry is the dialog's error and nothing changes); one hyperlink is set
//   over the selection's active range (an edited link keeps its own range
//   when the selection is its single cell); and an insert gives the linked
//   cells whose font colour is automatic the theme `hlink` colour and a
//   single underline (SPEC-123's applyStylePatches) in the same step. Remove
//   is one step "Remove link". Links read from files are never restyled.
// - The chip (LinkChip, `chip`): the active cell's link — stored, or a
//   top-level HYPERLINK() formula's, computed now — while no cell is edited
//   and the dialog is shut, until dismissed for that cell. Open follows it,
//   Edit opens the dialog (a formula's link: the formula in Edit mode), Copy
//   link copies its label, Remove removes a stored link.
// - Following (`follow`, the chip's Open and Ctrl/Cmd+click over a linked
//   cell's text — `pressFollows`, which SheetGrid asks before the pointer's
//   own handling, so it comes before SPEC-117's add-a-range; elsewhere in
//   the cell Ctrl/Cmd+click still adds a range): openHyperlinkTarget decides — external opens
//   window.open(href, "_blank", "noopener,noreferrer"); refused toasts
//   LINK_NOT_OPENABLE_MESSAGE; a place goes through the Name Box's reading
//   (interpretNameBox, the link's sheet the host) to activate its sheet,
//   select its range and reveal it — anything else (missing sheet, `#REF!`,
//   an unknown name) toasts LINK_DESTINATION_MISSING_MESSAGE, a hidden
//   sheet HIDDEN_SHEET_MESSAGE. A plain click only selects.
// - `isLinked(row, col)`: the sheet on show's cell carries data-link.
import { computed, nextTick, onScopeDispose, ref, shallowRef, type Ref, type ShallowRef } from "vue";
import { useToasts } from "../../composables/useToasts";
import { isMacPlatform } from "../../lib/core/platform";
import { findShortcut } from "../../lib/shortcuts";
import type { CellRef, GridRange } from "../../lib/spreadsheet/address";
import { externalHref, hyperlinkLabel, LINK_DESTINATION_MISSING_MESSAGE, LINK_NOT_OPENABLE_MESSAGE, LINK_UNDO_LABELS, openHyperlinkTarget, removeHyperlink, setHyperlink } from "../../lib/spreadsheet/annotations";
import { refitRows, type CellTextMeasurer } from "../../lib/spreadsheet/autofit";
import { setCellEntry } from "../../lib/spreadsheet/cells";
import { applyStylePatches } from "../../lib/spreadsheet/formatting";
import type { HyperlinkTarget, Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import { createCanvasCellMeasurer } from "./cellMeasurer";
import {
  displayTextAt,
  hyperlinkStyleEntries,
  linkAtCell,
  linkDialogFields,
  linkDialogResult,
  linkNameOptions,
  linkRange,
  linkSheetOptions,
  linkTargetAt,
  linkTextEntry,
  linkTextRef,
  type LinkDialogFields,
  type LinkOption,
} from "./cellLinks";
import { inGrid } from "./headerMenu";
import { HIDDEN_SHEET_MESSAGE, interpretNameBox } from "./nameBox";
import { SHEET_SHORTCUTS } from "./sheetShortcuts";
import type { CellEditing } from "./useCellEditing";
import { primaryModifier, type SheetSelectionController } from "./useSheetSelection";
import type { SheetViewState } from "./useSheetViewState";
import type { SpreadsheetUndo } from "./useSpreadsheetUndo";

export interface CellLinksDeps {
  book: {
    workbook: ShallowRef<SpreadsheetModel | null>;
    activeIndex: Ref<number>;
    activeSheet: Readonly<Ref<Sheet | null>>;
    selectSheet(index: number): void;
  };
  selection: Pick<SheetSelectionController, "selection" | "formatContext">;
  viewState: Pick<SheetViewState, "syncInto" | "update">;
  undo: Pick<SpreadsheetUndo, "record" | "commit">;
  /** Cell entry: busy while a cell is edited; focus and reveal after a jump. */
  cell: Pick<CellEditing, "editing" | "focusCell" | "revealTick" | "openEdit">;
  locale: () => SpreadsheetLocale;
  /** Shows a toast — default: useToasts' addToast. */
  toast?: (message: string) => void;
  /** Row refit's measurer — default: the canvas one. */
  measurer?: CellTextMeasurer;
  /** Opens an external link — default: window.open in a new, unlinked context. */
  openWindow?: (href: string) => void;
  /** Writes the chip's Copy link — default: the async clipboard. */
  copyText?: (text: string) => void;
}

/** The open dialog: its fields, its choices and what it edits. */
export interface LinkDialogState {
  fields: LinkDialogFields;
  /** Whether it edits the active cell's stored link (Remove shows). */
  editing: boolean;
  sheets: string[];
  names: LinkOption[];
  sheetId: string;
  at: CellRef;
  range: GridRange;
  linkId?: string;
  /** The stored target and the address it opened with (an unchanged address keeps it). */
  initial?: { target: HyperlinkTarget; address: string };
  /** The cell entry's refusal or formula error at the last Save. */
  entryError: string;
}

const single = (at: CellRef): GridRange => ({ top: at.row, left: at.col, bottom: at.row, right: at.col });

export function useCellLinks(deps: CellLinksDeps) {
  const toast = deps.toast ?? useToasts().addToast;
  const openWindow = deps.openWindow ?? ((href: string) => void window.open(href, "_blank", "noopener,noreferrer"));
  const copyText = deps.copyText ?? ((text: string) => void navigator.clipboard?.writeText(text).catch(() => undefined));
  let measurer = deps.measurer;
  const model = (): SpreadsheetModel | null => deps.book.workbook.value;
  const busy = (): boolean => deps.cell.editing.value;
  const dialog = shallowRef<LinkDialogState | null>(null);
  /** The cell whose chip was dismissed (`sheetId:row:col`). */
  const dismissed = ref("");

  /** The grid sheet on show, or null. */
  function gridSheet(): Sheet | null {
    const sheet = deps.book.activeSheet.value;
    return sheet?.kind === "grid" ? sheet : null;
  }

  /** Whether the sheet on show's cell is linked (data-link, the text's span). */
  function isLinked(row: number, col: number): boolean {
    const sheet = gridSheet();
    return sheet !== null && linkAtCell(sheet, { row, col }) !== null;
  }

  // --- the dialog --------------------------------------------------------------------------------

  /** Opens the dialog on the active cell (see the module header). */
  function openDialog(): void {
    const m = model();
    const sheet = gridSheet();
    if (!m || !sheet || busy()) {
      return;
    }
    const sel = deps.selection.selection.value;
    const at = sel.active;
    const active = sel.ranges[sel.activeRangeIndex] ?? single(at);
    const link = linkAtCell(sheet, at);
    const stored = link?.kind === "stored" ? link.link : null;
    const target = linkTargetAt(m, sheet, at);
    const fields = linkDialogFields(displayTextAt(m, sheet, at, deps.selection.formatContext.value), sheet.name, target, stored?.tooltip ?? "");
    dialog.value = {
      fields,
      editing: stored !== null,
      sheets: linkSheetOptions(m, fields.sheet),
      names: linkNameOptions(m),
      sheetId: sheet.id,
      at: { row: at.row, col: at.col },
      range: linkRange(active, stored),
      ...(stored ? { linkId: stored.linkId, initial: { target: stored.target, address: fields.address } } : {}),
      entryError: "",
    };
  }

  /** A field typed or chosen: a sheet or reference clears a chosen name. */
  function setField<K extends keyof LinkDialogFields>(key: K, value: LinkDialogFields[K]): void {
    const state = dialog.value;
    if (!state) {
      return;
    }
    const fields: LinkDialogFields = { ...state.fields, [key]: value };
    if (key === "sheet" || key === "ref") {
      fields.name = "";
    }
    dialog.value = { ...state, fields, entryError: "" };
  }

  const verdict = computed(() => (dialog.value ? linkDialogResult(dialog.value.fields) : null));
  /** The dialog's error line ("" = none). */
  const error = computed(() => {
    const result = verdict.value;
    return result && "error" in result ? result.error : (dialog.value?.entryError ?? "");
  });
  const canSave = computed(() => verdict.value !== null && !("error" in verdict.value));

  function close(): void {
    dialog.value = null;
    deps.cell.focusCell();
  }

  /** Save: ONE step (see the module header); false when nothing was saved. */
  function save(): boolean {
    const state = dialog.value;
    const result = verdict.value;
    const m = model();
    const sheetIndex = m?.sheets.findIndex((each) => each.id === state?.sheetId) ?? -1;
    const sheet = m?.sheets[sheetIndex];
    if (!state || !result || "error" in result || !m || !sheet) {
      return false;
    }
    const written = linkTextEntry(m, sheetIndex, state.at, result.text, deps.locale(), deps.selection.formatContext.value);
    if ("error" in written) {
      dialog.value = { ...state, entryError: written.error };
      return false;
    }
    const initial = state.initial;
    const target = initial && state.fields.kind === "web" && initial.target.kind === "url" && state.fields.address === initial.address ? initial.target : result.target;
    deps.viewState.syncInto(m);
    deps.undo.record(state.editing ? LINK_UNDO_LABELS.edit : LINK_UNDO_LABELS.insert, { sheetId: state.sheetId });
    if (written.entry) {
      const before = sheet.rows;
      setCellEntry(m, sheetIndex, state.at, written.entry);
      measurer ??= createCanvasCellMeasurer();
      refitRows(m, state.sheetId, [single(state.at)], measurer, deps.selection.formatContext.value, before);
    }
    setHyperlink(m, state.sheetId, state.range, target, result.tooltip);
    const linked = m.sheets[sheetIndex]!;
    const look = state.editing ? [] : hyperlinkStyleEntries(m, linked, state.range, state.at);
    if (look.length > 0) {
      applyStylePatches(m, state.sheetId, look);
    }
    deps.undo.commit();
    close();
    return true;
  }

  /** Removes a stored link: ONE step "Remove link". */
  function removeLink(sheetId: string, linkId: string): void {
    const m = model();
    if (!m) {
      return;
    }
    deps.viewState.syncInto(m);
    deps.undo.record(LINK_UNDO_LABELS.remove, { sheetId });
    removeHyperlink(m, sheetId, linkId);
    deps.undo.commit();
  }

  /** The dialog's Remove (an edited link only). */
  function remove(): void {
    const state = dialog.value;
    if (state?.linkId !== undefined) {
      removeLink(state.sheetId, state.linkId);
    }
    close();
  }

  // --- following ---------------------------------------------------------------------------------

  /** Selects `ranges` on `sheetId`, revealed and focused (the Name Box's jump). */
  async function jump(sheetId: string, ranges: GridRange[]): Promise<void> {
    const index = model()?.sheets.findIndex((sheet) => sheet.id === sheetId) ?? -1;
    const first = ranges[0];
    if (index < 0 || !first) {
      return;
    }
    if (index !== deps.book.activeIndex.value) {
      deps.book.selectSheet(index);
    }
    const active = { row: first.top, col: first.left };
    deps.viewState.update(sheetId, { selection: { active, ranges: [single(active)], activeRangeIndex: 0 } });
    await nextTick();
    deps.cell.revealTick.value += 1;
    await nextTick();
    deps.viewState.update(sheetId, { selection: { active, ranges: ranges.map((range) => ({ ...range })), activeRangeIndex: 0 } });
    deps.cell.focusCell();
  }

  /** Follows `target`, read on `host`'s sheet (see the module header). */
  async function follow(target: HyperlinkTarget, host: { sheetId: string; row: number; col: number }): Promise<void> {
    const m = model();
    const decision = openHyperlinkTarget(target);
    if (decision === "external" && target.kind === "url") {
      openWindow(externalHref(target));
    } else if (decision === "refused") {
      toast(LINK_NOT_OPENABLE_MESSAGE);
    } else if (m && target.kind === "place") {
      const outcome = interpretNameBox(m, target.location, host);
      if (outcome.kind === "jump") {
        await jump(outcome.sheetId, outcome.ranges);
      } else {
        toast(outcome.kind === "refuse" && outcome.message === HIDDEN_SHEET_MESSAGE ? HIDDEN_SHEET_MESSAGE : LINK_DESTINATION_MISSING_MESSAGE);
      }
    }
  }

  /** A grid press: Ctrl/Cmd (no Shift) over a linked cell's text follows
   * its link and takes the press (true); anything else is SPEC-117's. */
  function pressFollows(event: MouseEvent): boolean {
    const m = model();
    const sheet = gridSheet();
    const at = event.button === 0 && primaryModifier(event) && !event.shiftKey && !busy() ? linkTextRef(event.target) : null;
    const target = m && sheet && at ? linkTargetAt(m, sheet, at) : null;
    if (!sheet || !at || !target) {
      return false;
    }
    event.preventDefault();
    void follow(target, { sheetId: sheet.id, row: at.row, col: at.col });
    return true;
  }

  // --- the chip ----------------------------------------------------------------------------------

  /** The active cell's link and where it is (null: none). */
  const active = computed(() => {
    const m = model();
    const sheet = gridSheet();
    const at = deps.selection.selection.value.active;
    const link = sheet ? linkAtCell(sheet, at) : null;
    const target = m && sheet && link ? linkTargetAt(m, sheet, at) : null;
    return sheet && link && target ? { sheetId: sheet.id, at, link, target } : null;
  });
  const activeKey = computed(() => (active.value ? `${active.value.sheetId}:${active.value.at.row}:${active.value.at.col}` : ""));

  /** LinkChip's label and kind under the active cell (null: no chip). */
  const chip = computed(() => {
    const link = active.value;
    if (!link || busy() || dialog.value !== null || dismissed.value === activeKey.value) {
      return null;
    }
    return { label: hyperlinkLabel(link.target), external: link.target.kind === "url" };
  });

  function openFromChip(): void {
    const link = active.value;
    if (link) {
      void follow(link.target, { sheetId: link.sheetId, row: link.at.row, col: link.at.col });
    }
  }

  function editFromChip(): void {
    if (active.value?.link.kind === "formula") {
      deps.cell.openEdit();
    } else {
      openDialog();
    }
  }

  function copyFromChip(): void {
    if (active.value) {
      copyText(hyperlinkLabel(active.value.target));
    }
  }

  /** Removes a stored link (a formula's stays with its formula). */
  function removeFromChip(): void {
    const link = active.value;
    if (link?.link.kind === "stored") {
      removeLink(link.sheetId, link.link.link.linkId);
    } else {
      dismissed.value = activeKey.value;
    }
    deps.cell.focusCell();
  }

  function dismissChip(): void {
    dismissed.value = activeKey.value;
    deps.cell.focusCell();
  }

  // --- Ctrl/Cmd+K --------------------------------------------------------------------------------

  function onKeydown(event: KeyboardEvent): void {
    if (event.defaultPrevented || !inGrid(event.target) || busy()) {
      return;
    }
    if (findShortcut(event, isMacPlatform(), SHEET_SHORTCUTS)?.id === "sheet.insertLink") {
      event.preventDefault();
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

  return { dialog, error, canSave, openDialog, setField, save, remove, cancel: close, isLinked, follow, pressFollows, chip, openFromChip, editFromChip, copyFromChip, removeFromChip, dismissChip };
}

export type CellLinks = ReturnType<typeof useCellLinks>;

/** The grid context's links (SheetGrid's `context`), or null: not wired. */
export function linksOf(context: unknown): CellLinks | null {
  const value = context !== null && typeof context === "object" ? Reflect.get(context, "links") : undefined;
  return value ? (value as CellLinks) : null;
}

/** Narrows a component's runtime-typed `links` prop (the .vue parser bars casts). */
export function asCellLinks(value: unknown): CellLinks {
  return value as CellLinks;
}
