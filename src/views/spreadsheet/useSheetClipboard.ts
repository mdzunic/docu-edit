// Cut, copy and paste in the workbook view (SPEC-125 "Flavours", "Paste
// geometry", "Paste semantics", "Cut and move", "Shortcuts") over
// lib/spreadsheet/clipboard.ts, paste.ts and move.ts — the spreadsheet
// sibling of the text product's useClipboard (SPEC-016).
//
// - Copy and Cut are the native `copy`/`cut` events while the grid has focus
//   and no cell is being edited (the window's capture phase, so the in-cell
//   editor holding the grid's focus never sees them): both flavours are
//   written in the synchronous handler with clipboardData.setData, a Cut
//   with a fresh cut token, and the model is untouched. A Copy over
//   filter-hidden rows takes the visible rows only (SPEC-128, clipboard.ts);
//   a Cut takes every row, as the move it pastes as does. A multi-range
//   selection or more than MAX_COPY_CELLS refuses with a toast. The copied
//   range is outlined by the marching ants (`marquee`, view state — never in
//   the undo stack) until Escape, another copy or cut, any edit (the undo
//   stack's revision, but for this view's own paste of a copy), or the paste
//   of a cut.
// - Paste (Ctrl/Cmd+V in the grid; in Edit mode the cell editor's own
//   plain-text paste, SPEC-122) reads text/html first — a
//   table[data-docu-edit-cells] whose payload validates (this view's own
//   copy: its in-memory payload and style ids), else an external table
//   (clipboardImport.ts) — then text/plain. A payload carrying this view's
//   pending cut token, from this workbook, moves the cut cells (moveRange,
//   "Move cells"; a cut across sheets records the workbook) — once; anything
//   else pastes as a copy ("Paste"), then selects what it pasted.
// - The cell menu's Copy and Cut fire the same events through
//   document.execCommand, as the text product's menu does; its Paste reads
//   the async API (sheetClipboardIo.ts) — HTML without the attribute (or no
//   HTML) whose text/plain equals this view's last copy's TSV is that copy
//   (SPEC-016's snapshot rule) — and an unavailable clipboard toasts
//   SPEC-016's message.
// - Paste special (Ctrl/Cmd+Shift+V, Ctrl+Alt+V, the Edit and cell menus)
//   reads the async API — refused: CLIPBOARD_BLOCKED_MESSAGE — and opens
//   PasteSpecialDialog with what the source can supply ("Paste special").
// - Every paste is ONE undo step with the rows it wrote refitted
//   (sheetSteps.ts); a refusal toasts and records nothing. SPEC-131: a pasted
//   image (sheetClipboardIo's pastedImage) inserts as a drawing instead.
import { onScopeDispose, shallowRef, watch, type Ref, type ShallowRef } from "vue";
import { useToasts } from "../../composables/useToasts";
import { isMacPlatform } from "../../lib/core/platform";
import { findShortcut } from "../../lib/shortcuts";
import type { GridRange } from "../../lib/spreadsheet/address";
import type { CellTextMeasurer } from "../../lib/spreadsheet/autofit";
import {
  buildClipboardCopy,
  CLIPBOARD_BLOCKED_MESSAGE,
  CLIPBOARD_UNAVAILABLE_MESSAGE,
  copyFlavours,
  copyRefusal,
  formulasAsValuesMessage,
  MULTIPLE_SELECTIONS_MESSAGE,
  parseClipboardPayload,
  type ClipboardCopy,
} from "../../lib/spreadsheet/clipboard";
import { importClipboardHtml } from "../../lib/spreadsheet/clipboardImport";
import type { Sheet, SpreadsheetModel } from "../../lib/spreadsheet/model";
import { moveRange, type MoveResult } from "../../lib/spreadsheet/move";
import type { SpreadsheetLocale } from "../../lib/spreadsheet/numberFormat/locale";
import { blockFromPayload, blockFromText, pasteBlock, pasteCapabilities, type PasteBlock, type PasteCapabilities, type PasteOptions, type PasteResult } from "../../lib/spreadsheet/paste";
import { canonicalFormula } from "./cellEntry";
import { createCanvasCellMeasurer } from "./cellMeasurer";
import { inGrid } from "./headerMenu";
import { newCutToken, pastedImage, payloadAttribute, readClipboard } from "./sheetClipboardIo";
import { SHEET_SHORTCUTS } from "./sheetShortcuts";
import { runStep, selectionOf, type StepContext } from "./sheetSteps";
import type { CellEditing } from "./useCellEditing";
import type { SheetDrawings } from "./useSheetDrawings";
import type { SheetSelectionController } from "./useSheetSelection";
import type { SheetViewState } from "./useSheetViewState";
import type { SpreadsheetUndo } from "./useSpreadsheetUndo";

export interface SheetClipboardDeps {
  book: { workbook: ShallowRef<SpreadsheetModel | null>; activeSheet: Readonly<Ref<Sheet | null>> };
  viewState: Pick<SheetViewState, "syncInto">;
  selection: Pick<SheetSelectionController, "selection" | "set" | "formatContext">;
  editing: { undo: Pick<SpreadsheetUndo, "record" | "commit" | "revision">; cell: Pick<CellEditing, "editing" | "focusCell">; drawings?: Pick<SheetDrawings, "insertImageFile"> };
  locale: () => SpreadsheetLocale;
  /** Shows a toast — default: useToasts' addToast. */
  toast?: (message: string) => void;
  /** Row refit's measurer — default: the canvas one. */
  measurer?: CellTextMeasurer;
  /** The async Clipboard API — default: navigator.clipboard. */
  clipboard?: () => Clipboard | undefined;
}

export type ClipboardMenuAction = "cut" | "copy" | "paste" | "paste-special";

/** A block about to be pasted, with the source style ids of this view's own
 * copy (a same-workbook paste keeps them by identity); `own`: it is this
 * view's last copy. */
interface Source {
  block: PasteBlock;
  styleIds?: readonly number[];
  own?: true;
}

export function useSheetClipboard(deps: SheetClipboardDeps) {
  const toast = deps.toast ?? useToasts().addToast;
  let measurer = deps.measurer;
  /** The marching ants: the copied range, on its sheet. */
  const marquee = shallowRef<{ sheetId: string; range: GridRange; mode: "copy" | "cut" } | null>(null);
  /** The open Paste special dialog: its source and what it can supply. */
  const pasteSpecial = shallowRef<(Source & { capabilities: PasteCapabilities }) | null>(null);
  let lastCopy: { copy: ClipboardCopy; text: string; attribute: string } | null = null;
  let pendingCut: { token: string; sheetId: string; range: GridRange } | null = null;
  /** The undo revision the ants were drawn at, or kept at by this view's own
   * paste of its copy: any other revision is an edit made since. */
  let antsRevision = -1;

  const model = (): SpreadsheetModel | null => deps.book.workbook.value;
  const gridSheet = (): Sheet | null => {
    const sheet = deps.book.activeSheet.value;
    return sheet && sheet.kind === "grid" ? sheet : null;
  };
  const ready = (): boolean => model() !== null && gridSheet() !== null && !deps.editing.cell.editing.value;

  function clearAnts(): void {
    marquee.value = null;
    pendingCut = null;
  }

  // Any edit ends the ants and cancels a pending cut — but this view's own
  // paste of a copy (Excel keeps the ants so the copy pastes again). The
  // watcher runs after the edit's task, so it compares against the revision
  // the ants were drawn at: a copy made after an edit keeps its ants.
  watch(
    () => deps.editing.undo.revision.value,
    (revision) => {
      if (revision !== antsRevision) {
        clearAnts();
      }
    },
  );

  function context(m: SpreadsheetModel): StepContext {
    measurer ??= createCanvasCellMeasurer();
    return { model: m, undo: deps.editing.undo, measurer, format: deps.selection.formatContext.value, sync: (each) => deps.viewState.syncInto(each) };
  }

  /** Copy (or Cut): both flavours into the event, the ants on. */
  function writeCopy(event: ClipboardEvent, cut: boolean): void {
    const m = model();
    const sheet = gridSheet();
    const data = event.clipboardData;
    if (!m || !sheet || !data || !ready() || !inGrid(event.target)) {
      return;
    }
    const ranges = deps.selection.selection.value.ranges;
    const refusal = copyRefusal(sheet, ranges);
    if (refusal) {
      toast(refusal);
      return;
    }
    event.preventDefault();
    const token = cut ? newCutToken() : undefined;
    const copy = buildClipboardCopy(m, sheet.id, ranges[0]!, token, { visibleRows: true });
    const { text, html, attribute } = copyFlavours(m, sheet.id, copy, deps.selection.formatContext.value);
    data.setData("text/plain", text);
    data.setData("text/html", html);
    lastCopy = { copy, text, attribute };
    marquee.value = { sheetId: sheet.id, range: copy.area.range, mode: cut ? "cut" : "copy" };
    pendingCut = token ? { token, sheetId: sheet.id, range: copy.area.range } : null;
    antsRevision = deps.editing.undo.revision.value;
  }

  /** What the clipboard's flavours hold (see the module header): a source, a
   * refusal, or null for nothing to paste. `viaMenu`: the async read, where
   * SPEC-016's snapshot rule applies. */
  function resolve(html: string, text: string, viaMenu: boolean): Source | string | null {
    const m = model();
    if (!m) {
      return null;
    }
    const own: Source | null = lastCopy ? { block: blockFromPayload(lastCopy.copy.payload), styleIds: lastCopy.copy.styleIds, own: true } : null;
    const ctx = { locale: deps.locale(), dateSystem: m.dateSystem, canonical: canonicalFormula };
    const attribute = html ? payloadAttribute(html) : null;
    if (attribute !== null) {
      if (own && attribute === lastCopy?.attribute) {
        return own;
      }
      const payload = parseClipboardPayload(attribute);
      if (payload) {
        return { block: blockFromPayload(payload) };
      }
    } else if (viaMenu && own && text !== "" && text === lastCopy?.text) {
      // The snapshot rule: HTML without the attribute — or no HTML at all,
      // as a read may offer — whose text is the last copy's is that copy.
      return own;
    }
    if (html) {
      const imported = importClipboardHtml(html, ctx);
      if (imported !== null) {
        return typeof imported === "string" ? imported : { block: imported };
      }
    }
    const fromText = blockFromText(text, ctx);
    return fromText === null || typeof fromText === "string" ? fromText : { block: fromText };
  }

  /** A cut pasted where it came from: the cells move (see "Cut and move"). */
  function pasteCut(m: SpreadsheetModel, sheet: Sheet, cut: NonNullable<typeof pendingCut>): void {
    const ranges = deps.selection.selection.value.ranges;
    if (ranges.length !== 1) {
      toast(MULTIPLE_SELECTIONS_MESSAGE);
      return;
    }
    const to = ranges[0]!;
    const move = { fromSheetId: cut.sheetId, from: cut.range, toSheetId: sheet.id, toTop: to.top, toLeft: to.left };
    const dest = { top: to.top, left: to.left, bottom: to.top + cut.range.bottom - cut.range.top, right: to.left + cut.range.right - cut.range.left };
    const scope = cut.sheetId === sheet.id ? { sheetId: sheet.id } : "workbook";
    const result: MoveResult = runStep(context(m), "Move cells", scope, () => moveRange(m, move), () => [
      { sheetId: sheet.id, ranges: [dest] },
      { sheetId: cut.sheetId, ranges: [cut.range] },
    ]);
    if (!result.ok) {
      toast(result.message);
      return;
    }
    clearAnts();
    deps.selection.set(selectionOf(dest));
  }

  /** Pastes a source onto the selection as one step `label`. */
  function paste(source: Source, options: PasteOptions, label: string): void {
    const m = model();
    const sheet = gridSheet();
    if (!m || !sheet || !ready()) {
      return;
    }
    const cut = pendingCut;
    if (label === "Paste" && cut && source.block.cut === cut.token && source.block.workbookId === m.id && m.sheets.some((each) => each.id === cut.sheetId)) {
      pasteCut(m, sheet, cut);
      return;
    }
    const selection = deps.selection.selection.value;
    const result: PasteResult = runStep(
      context(m),
      label,
      { sheetId: sheet.id },
      () => pasteBlock(m, sheet.id, { ranges: selection.ranges, active: selection.active }, source.block, options, source.styleIds ? { styleIds: source.styleIds } : {}),
      (done) => (done.ok ? [{ sheetId: sheet.id, ranges: [done.range] }] : []),
    );
    if (!result.ok) {
      toast(result.message);
      return;
    }
    if (source.own && source.block.cut === undefined) {
      // This view's own copy pasted: the ants stay, so it pastes again.
      // Any other paste is an edit like the rest — it ends the ants and
      // cancels a pending cut (the revision watcher above).
      antsRevision = deps.editing.undo.revision.value;
    }
    deps.selection.set(selectionOf(result.range));
    if (result.formulasAsValues > 0) {
      toast(formulasAsValuesMessage(result.formulasAsValues));
    }
  }

  function onCopy(event: ClipboardEvent): void {
    writeCopy(event, false);
  }

  function onCut(event: ClipboardEvent): void {
    writeCopy(event, true);
  }

  function onPaste(event: ClipboardEvent): void {
    const data = event.clipboardData;
    if (!data || !ready() || !inGrid(event.target)) {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    const image = deps.editing.drawings ? pastedImage(data) : null;
    if (image) {
      void deps.editing.drawings?.insertImageFile(image);
      return;
    }
    const source = resolve(data.getData("text/html"), data.getData("text/plain"), false);
    if (typeof source === "string") {
      toast(source);
    } else if (source) {
      paste(source, { what: "all" }, "Paste");
    }
  }

  /** The cell menu's Paste: the async read (see the module header). */
  async function menuPaste(): Promise<void> {
    const flavours = await readClipboard(deps.clipboard?.() ?? globalThis.navigator?.clipboard);
    if (!flavours) {
      toast(CLIPBOARD_UNAVAILABLE_MESSAGE);
      return;
    }
    const source = resolve(flavours.html, flavours.text, true);
    if (typeof source === "string") {
      toast(source);
    } else if (source) {
      paste(source, { what: "all" }, "Paste");
    }
  }

  /** Paste special: reads the clipboard and opens the dialog. */
  async function openPasteSpecial(): Promise<void> {
    if (!ready()) {
      return;
    }
    const flavours = await readClipboard(deps.clipboard?.() ?? globalThis.navigator?.clipboard);
    if (!flavours) {
      toast(CLIPBOARD_BLOCKED_MESSAGE);
      return;
    }
    const source = resolve(flavours.html, flavours.text, true);
    if (typeof source === "string") {
      toast(source);
    } else if (source) {
      pasteSpecial.value = { ...source, capabilities: pasteCapabilities(source.block) };
    }
  }

  /** The cell menu's Cut, Copy, Paste and Paste special… */
  function menu(action: ClipboardMenuAction): void {
    if (action === "copy" || action === "cut") {
      deps.editing.cell.focusCell();
      if (typeof document.execCommand === "function") {
        document.execCommand(action);
      }
    } else if (action === "paste") {
      void menuPaste();
    } else {
      void openPasteSpecial();
    }
  }

  function onKeydown(event: KeyboardEvent): void {
    if (!ready() || !inGrid(event.target)) {
      return;
    }
    if (event.key === "Escape" && marquee.value) {
      clearAnts();
      return;
    }
    if (findShortcut(event, isMacPlatform(), SHEET_SHORTCUTS)?.id === "sheet.pasteSpecial") {
      event.preventDefault();
      void openPasteSpecial();
    }
  }

  if (typeof window !== "undefined") {
    window.addEventListener("copy", onCopy, true);
    window.addEventListener("cut", onCut, true);
    window.addEventListener("paste", onPaste, true);
    window.addEventListener("keydown", onKeydown, true);
  }
  onScopeDispose(() => {
    if (typeof window !== "undefined") {
      window.removeEventListener("copy", onCopy, true);
      window.removeEventListener("cut", onCut, true);
      window.removeEventListener("paste", onPaste, true);
      window.removeEventListener("keydown", onKeydown, true);
    }
  });

  return {
    marquee,
    pasteSpecial,
    menu,
    openPasteSpecial,
    /** The dialog's OK: one "Paste special" step (a cut pastes as a copy). */
    confirmPasteSpecial(options: PasteOptions): void {
      const source = pasteSpecial.value;
      pasteSpecial.value = null;
      if (source) {
        paste(source, options, "Paste special");
      }
    },
    cancelPasteSpecial(): void {
      pasteSpecial.value = null;
    },
    /** Escape's work: the ants off and a pending cut cancelled. */
    clearAnts,
    onCopy,
    onCut,
    onPaste,
    onKeydown,
  };
}

export type SheetClipboard = ReturnType<typeof useSheetClipboard>;

/** Narrows a component's runtime-typed `clipboard` prop (the .vue parser
 * bars casts). */
export function asSheetClipboard(value: unknown): SheetClipboard {
  return value as SheetClipboard;
}
