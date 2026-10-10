// The sheet's keyboard shortcuts (SPEC-123 "Shortcuts"): the formatting
// commands the spreadsheet grid binds, declared in SPEC-021's
// ShortcutDefinition shape and matched and labelled only through
// src/lib/shortcuts.ts — findShortcut(event, isMac, SHEET_SHORTCUTS),
// shortcutLabel(id, isMac, SHEET_SHORTCUTS). A table of its own rather than
// entries in SHORTCUTS: Ctrl/Cmd+B/I/U are already the text editor's
// format.* there, and the text registry's matching and reference dialog stay
// exactly as they were.
//
// The number formats bind physical keys (`code`, as Excel binds them), with
// `key` only the character shown, so they are the same keys on every layout:
// on a German keyboard, where Shift+3 types "§", Ctrl+Shift+3 is still the
// date format. macOS binds them to Control, not Cmd, because Cmd+Shift+3/4/5
// are the system's screenshot keys. Strikethrough is the one binding that
// differs by platform (Ctrl+5 on Windows/Linux, Cmd+Shift+X on macOS), hence
// sheetShortcutDefinitions(isMac). No two bindings may collide on either
// platform and none may claim RESERVED_BROWSER_KEYS; shortcuts.spec.ts
// asserts both.
//
// The bindings act only when focus is in the grid — not the cell editor while
// editing, the formula bar or a dialog; the caller enforces that. Chrome and
// Firefox keep Ctrl/Cmd+1 … 8 for switching browser tabs, so Ctrl/Cmd+1
// (Format cells) and Windows/Linux Ctrl+5 (strikethrough) are best effort:
// the Format menu and the toolbar's strikethrough button
// (`sheet-strikethrough`) are the guaranteed paths.
//
// SPEC-124 adds the rows-and-columns pair, Google Sheets' binding: Ctrl+Alt+=
// / Cmd+Option+= inserts (`sheet.insertCells`) and Ctrl+Alt+- /
// Cmd+Option+- deletes (`sheet.deleteCells`), by physical key (`Equal`,
// `Minus`, so Option's dead characters on macOS do not matter). Ctrl/Cmd+=,
// Ctrl/Cmd+- and Ctrl/Cmd+Shift+= stay the sheet zoom (SPEC-063's
// view.zoomIn/zoomOut).
//
// SPEC-125 adds paste special — Ctrl/Cmd+Shift+V, alternate Ctrl+Alt+V
// (Control+Option+V on macOS), both by physical key (`KeyV`) — and fill down
// / right, Ctrl/Cmd+D and Ctrl/Cmd+R. D and R are the browser's bookmark and
// reload keys: useFillHandle claims them only while the grid has focus (as
// Google Sheets does); Edit ▸ Fill ▸ Down/Right is the menu path. Copy, cut
// and paste are the native clipboard events, not entries here.
import { isMacPlatform } from "../../lib/core/platform";
import type { ShortcutDefinition } from "../../lib/shortcuts";
import type { NumberShortcutId } from "../../lib/spreadsheet/numberCodes";

export type SheetShortcutId =
  | "sheet.formatCells"
  | "sheet.bold"
  | "sheet.italic"
  | "sheet.underline"
  | "sheet.strikethrough"
  | "sheet.formatGeneral"
  | "sheet.formatNumber"
  | "sheet.formatTime"
  | "sheet.formatDate"
  | "sheet.formatCurrency"
  | "sheet.formatPercent"
  | "sheet.formatScientific"
  | "sheet.insertCells"
  | "sheet.deleteCells"
  | "sheet.pasteSpecial"
  | "sheet.fillDown"
  | "sheet.fillRight";

const CATEGORY = "Sheet formatting";
const STRUCTURE_CATEGORY = "Rows and columns";
const PASTE_CATEGORY = "Paste and fill";

/** The sheet's table for one platform, one binding per command. Bold, italic
 * and underline override the same browser defaults the text editor's do
 * (SPEC-021 AC-10). */
export function sheetShortcutDefinitions(isMac: boolean): ShortcutDefinition[] {
  const definitions: (ShortcutDefinition & { id: SheetShortcutId })[] = [
    { id: "sheet.formatCells", label: "Format cells", category: CATEGORY, bindings: [{ key: "1", mod: true }] },
    { id: "sheet.bold", label: "Bold", category: CATEGORY, bindings: [{ key: "b", mod: true }], overridesBrowser: true },
    { id: "sheet.italic", label: "Italic", category: CATEGORY, bindings: [{ key: "i", mod: true }], overridesBrowser: true },
    { id: "sheet.underline", label: "Underline", category: CATEGORY, bindings: [{ key: "u", mod: true }], overridesBrowser: true },
    {
      id: "sheet.strikethrough",
      label: "Strikethrough",
      category: CATEGORY,
      bindings: [isMac ? { key: "x", mod: true, shift: true } : { key: "5", mod: true }],
    },
    // Ctrl (macOS: Control) + Shift + a physical key; `key` is what a US
    // layout types there, shown in labels.
    { id: "sheet.formatGeneral", label: "General number format", category: CATEGORY, bindings: [{ key: "~", code: "Backquote", ctrl: true, shift: true }] },
    { id: "sheet.formatNumber", label: "Number format", category: CATEGORY, bindings: [{ key: "!", code: "Digit1", ctrl: true, shift: true }] },
    { id: "sheet.formatTime", label: "Time format", category: CATEGORY, bindings: [{ key: "@", code: "Digit2", ctrl: true, shift: true }] },
    { id: "sheet.formatDate", label: "Date format", category: CATEGORY, bindings: [{ key: "#", code: "Digit3", ctrl: true, shift: true }] },
    { id: "sheet.formatCurrency", label: "Currency format", category: CATEGORY, bindings: [{ key: "$", code: "Digit4", ctrl: true, shift: true }] },
    { id: "sheet.formatPercent", label: "Percent format", category: CATEGORY, bindings: [{ key: "%", code: "Digit5", ctrl: true, shift: true }] },
    { id: "sheet.formatScientific", label: "Scientific format", category: CATEGORY, bindings: [{ key: "^", code: "Digit6", ctrl: true, shift: true }] },
    // SPEC-124: Ctrl/Cmd + Alt/Option + a physical key.
    { id: "sheet.insertCells", label: "Insert rows or columns", category: STRUCTURE_CATEGORY, bindings: [{ key: "=", code: "Equal", mod: true, alt: true }] },
    { id: "sheet.deleteCells", label: "Delete rows or columns", category: STRUCTURE_CATEGORY, bindings: [{ key: "-", code: "Minus", mod: true, alt: true }] },
    // SPEC-125: paste special by physical key, fill down and right.
    {
      id: "sheet.pasteSpecial",
      label: "Paste special",
      category: PASTE_CATEGORY,
      bindings: [
        { key: "V", code: "KeyV", mod: true, shift: true },
        { key: "V", code: "KeyV", ctrl: true, alt: true },
      ],
    },
    { id: "sheet.fillDown", label: "Fill down", category: PASTE_CATEGORY, bindings: [{ key: "d", mod: true }] },
    { id: "sheet.fillRight", label: "Fill right", category: PASTE_CATEGORY, bindings: [{ key: "r", mod: true }] },
  ];
  return definitions;
}

/** The current platform's table, resolved once at load. Matching or labelling
 * for another platform (a test, a preview) takes sheetShortcutDefinitions
 * with that same `isMac` instead — the strikethrough binding differs. */
export const SHEET_SHORTCUTS: ShortcutDefinition[] = sheetShortcutDefinitions(isMacPlatform());

/** Whether a shortcut is one of the formatting commands useFormatting runs
 * (the others are useStructure's, useSheetClipboard's and useFillHandle's). */
export function isFormattingShortcut(id: string): boolean {
  return SHEET_SHORTCUTS.find((definition) => definition.id === id)?.category === CATEGORY;
}

/** The number-format shortcut ids → their format (useFormatting runs them). */
export const NUMBER_SHORTCUTS: Readonly<Record<string, NumberShortcutId>> = {
  "sheet.formatGeneral": "general",
  "sheet.formatNumber": "number",
  "sheet.formatTime": "time",
  "sheet.formatDate": "date",
  "sheet.formatCurrency": "currency",
  "sheet.formatPercent": "percent",
  "sheet.formatScientific": "scientific",
};

/** The toggle shortcut ids → their format, and the toggles' undo labels. */
export const TOGGLE_SHORTCUTS: Readonly<Record<string, "bold" | "italic" | "underline" | "strikethrough">> = {
  "sheet.bold": "bold",
  "sheet.italic": "italic",
  "sheet.underline": "underline",
  "sheet.strikethrough": "strikethrough",
};
export const TOGGLE_LABELS = { bold: "Bold", italic: "Italic", underline: "Underline", strikethrough: "Strikethrough" } as const;
