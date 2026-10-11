// The single keyboard-shortcut registry (SPEC-021). Every key binding the
// app intercepts is declared here — command id, display label, category and
// platform-neutral binding(s) — so the editor's keydown handlers, toolbar
// tooltips and the "Keyboard shortcuts" reference dialog all read the same
// source of truth (AC-1, AC-2). `mod` means "Ctrl on Windows/Linux, Cmd on
// macOS", resolved at runtime via `isMacPlatform()` — never by sniffing the
// user-agent string (AC-3, AC-4).
//
// SPEC-123: a product with commands of its own declares them beside its view
// in this same shape — the sheet's SHEET_SHORTCUTS — and matches and labels
// them here through findShortcut/shortcutLabel's `definitions` parameter,
// which defaults to this registry. Bindings gain `code`, a physical key
// (`KeyboardEvent.code`) so a command holds on every keyboard layout, and
// `ctrl`, the Control key on every platform (on macOS Control itself, not
// Cmd). A binding with neither matches exactly as it always has.
//
// Deliberately excluded: per-component dismissal conventions (Escape closes
// a menu/popover/dialog, Enter confirms a text field) aren't registered
// here. Those are a generic, universal UI convention every overlay in the
// app implements locally, not a named, documented command — real word
// processors don't list "Escape closes this dialog" in their shortcut
// reference either. The registry is for commands a user would look up by
// name.
import { isMacPlatform } from "./core/platform";

/** One key combination that triggers a command. `key` is compared against
 * `KeyboardEvent.key` — case-insensitively for single characters, exactly
 * for named keys ("Home", "Tab", "PageUp", "PageDown"). Absence of `mod`/
 * `shift`/`alt` means that modifier must NOT be held — bindings are matched
 * exactly, not as a subset, so "Ctrl+Z" and "Ctrl+Shift+Z" never collide.
 * `shiftOptional` is for the handful of commands (page/document navigation)
 * where Shift doesn't select a *different* command, only how this one
 * behaves — Ctrl+End moves the caret, Ctrl+Shift+End extends the selection
 * to the same target; both are "nav.documentEnd", and the handler reads
 * `event.shiftKey` itself for the extend/move distinction. */
export interface ShortcutBinding {
  key: string;
  /** SPEC-123: a physical key (`KeyboardEvent.code`, e.g. "Digit1",
   * "Backquote") matched instead of `key`, which then is only the display
   * character ("!", "~") — so the binding is the same key position on every
   * layout, as Excel binds its Ctrl+Shift+digit number formats. */
  code?: string;
  mod?: boolean;
  /** SPEC-123: the Control key on every platform. On macOS that is Control
   * itself, with Cmd held exactly when `mod` is also set; on Windows/Linux
   * `mod` already is Ctrl, so `ctrl` satisfies it. */
  ctrl?: boolean;
  shift?: boolean;
  /** Alt on Windows/Linux, Option on macOS (SPEC-019's style shortcuts). */
  alt?: boolean;
  shiftOptional?: boolean;
}

export interface ShortcutDefinition {
  /** Stable id used for dispatch, e.g. "format.bold" — never shown to the user. */
  id: string;
  label: string;
  category: string;
  /** Every combination that triggers this command. The first entry is the
   * one shown in tooltips and used as the display binding; later entries are
   * alternates (e.g. redo's Ctrl+Y). */
  bindings: ShortcutBinding[];
  /** True for the five browser defaults this app deliberately overrides —
   * Ctrl/Cmd+B/I/U/F/P (AC-10). Shown with a note in the reference dialog. */
  overridesBrowser?: boolean;
}

/** Ctrl/Cmd+T/W/N (new tab, close tab, new window) must never be claimed by
 * the app (AC-11) — asserted against this registry, and SPEC-123's
 * SHEET_SHORTCUTS, in shortcuts.spec.ts. */
export const RESERVED_BROWSER_KEYS = ["t", "w", "n"];

export const SHORTCUTS: ShortcutDefinition[] = [
  // Formatting
  { id: "format.bold", label: "Bold", category: "Formatting", bindings: [{ key: "b", mod: true }], overridesBrowser: true },
  { id: "format.italic", label: "Italic", category: "Formatting", bindings: [{ key: "i", mod: true }], overridesBrowser: true },
  { id: "format.underline", label: "Underline", category: "Formatting", bindings: [{ key: "u", mod: true }], overridesBrowser: true },
  { id: "format.strikethrough", label: "Strikethrough", category: "Formatting", bindings: [{ key: "x", mod: true, shift: true }] },
  // SPEC-066 claims the conventional Cmd/Ctrl+, for Preferences (AC-2), the
  // macOS system convention. Subscript moves to the same physical comma key
  // with Shift, which the browser reports as "<" (bindings match the produced
  // character, like view.zoomIn's "+" alternate) — so no binding collides.
  { id: "format.subscript", label: "Subscript", category: "Formatting", bindings: [{ key: "<", mod: true, shift: true }] },
  { id: "format.superscript", label: "Superscript", category: "Formatting", bindings: [{ key: ".", mod: true }] },
  { id: "format.clearFormatting", label: "Clear formatting", category: "Formatting", bindings: [{ key: "\\", mod: true }] },

  // Paragraph
  { id: "paragraph.alignLeft", label: "Align left", category: "Paragraph", bindings: [{ key: "l", mod: true }] },
  { id: "paragraph.alignCenter", label: "Align centre", category: "Paragraph", bindings: [{ key: "e", mod: true }] },
  { id: "paragraph.alignRight", label: "Align right", category: "Paragraph", bindings: [{ key: "r", mod: true }] },
  { id: "paragraph.alignJustify", label: "Justify", category: "Paragraph", bindings: [{ key: "j", mod: true }] },
  { id: "paragraph.indent", label: "Increase indent / demote list item", category: "Paragraph", bindings: [{ key: "Tab" }] },
  { id: "paragraph.outdent", label: "Decrease indent / promote list item", category: "Paragraph", bindings: [{ key: "Tab", shift: true }] },

  // Paragraph styles (SPEC-019 AC-5, AC-6)
  { id: "style.normal", label: "Apply Normal text", category: "Paragraph styles", bindings: [{ key: "0", mod: true, alt: true }] },
  { id: "style.heading1", label: "Apply Heading 1", category: "Paragraph styles", bindings: [{ key: "1", mod: true, alt: true }] },
  { id: "style.heading2", label: "Apply Heading 2", category: "Paragraph styles", bindings: [{ key: "2", mod: true, alt: true }] },
  { id: "style.heading3", label: "Apply Heading 3", category: "Paragraph styles", bindings: [{ key: "3", mod: true, alt: true }] },
  { id: "style.heading4", label: "Apply Heading 4", category: "Paragraph styles", bindings: [{ key: "4", mod: true, alt: true }] },
  { id: "style.heading5", label: "Apply Heading 5", category: "Paragraph styles", bindings: [{ key: "5", mod: true, alt: true }] },
  { id: "style.heading6", label: "Apply Heading 6", category: "Paragraph styles", bindings: [{ key: "6", mod: true, alt: true }] },

  // Editing
  { id: "edit.undo", label: "Undo", category: "Editing", bindings: [{ key: "z", mod: true }] },
  {
    id: "edit.redo",
    label: "Redo",
    category: "Editing",
    bindings: [
      { key: "z", mod: true, shift: true },
      { key: "y", mod: true },
    ],
  },
  { id: "edit.selectAll", label: "Select all", category: "Editing", bindings: [{ key: "a", mod: true }] },
  { id: "edit.pastePlain", label: "Paste without formatting", category: "Editing", bindings: [{ key: "v", mod: true, shift: true }] },
  // SPEC-025 first bound Insert link to Ctrl/Cmd+K. SPEC-062 reclaims that key
  // for the command palette (its AC-1), so the link accelerator moves to
  // Ctrl/Cmd+Alt+K — joining the insert.* Alt family (footnote ⌥F, endnote ⌥D,
  // comment ⌥M). Only the accelerator changes; the command, handler and dialog
  // are untouched.
  { id: "insert.link", label: "Insert link", category: "Editing", bindings: [{ key: "k", mod: true, alt: true }] },

  // SPEC-053: F9 updates the field at the caret.
  { id: "field.update", label: "Update field", category: "Editing", bindings: [{ key: "F9" }] },

  // Find & replace
  { id: "find.open", label: "Find", category: "Find & replace", bindings: [{ key: "f", mod: true }], overridesBrowser: true },
  { id: "find.openReplace", label: "Find and replace", category: "Find & replace", bindings: [{ key: "h", mod: true }] },

  // Navigation
  {
    id: "nav.documentStart",
    label: "Go to start of document",
    category: "Navigation",
    bindings: [{ key: "Home", mod: true, shiftOptional: true }],
  },
  {
    id: "nav.documentEnd",
    label: "Go to end of document",
    category: "Navigation",
    bindings: [{ key: "End", mod: true, shiftOptional: true }],
  },
  { id: "nav.pageUp", label: "Move up one page", category: "Navigation", bindings: [{ key: "PageUp", shiftOptional: true }] },
  {
    id: "nav.pageDown",
    label: "Move down one page",
    category: "Navigation",
    bindings: [{ key: "PageDown", shiftOptional: true }],
  },

  // Insert (SPEC-026; page break SPEC-028)
  {
    id: "insert.nonBreakingSpace",
    label: "Insert non-breaking space",
    category: "Insert",
    bindings: [{ key: " ", mod: true, shift: true }],
  },
  {
    id: "insert.pageBreak",
    label: "Insert page break",
    category: "Insert",
    bindings: [{ key: "Enter", mod: true }],
  },
  {
    id: "insert.equation",
    label: "Insert equation",
    category: "Insert",
    bindings: [{ key: "=", alt: true }],
  },
  {
    id: "insert.footnote",
    label: "Insert footnote",
    category: "Insert",
    bindings: [{ key: "f", mod: true, alt: true }],
  },
  {
    id: "insert.endnote",
    label: "Insert endnote",
    category: "Insert",
    bindings: [{ key: "d", mod: true, alt: true }],
  },
  {
    id: "insert.comment",
    label: "Insert comment",
    category: "Insert",
    bindings: [{ key: "m", mod: true, alt: true }],
  },

  // Editing
  {
    id: "edit.toggleTrackChanges",
    label: "Track changes",
    category: "Editing",
    bindings: [{ key: "e", mod: true, shift: true }],
  },

  // File
  { id: "file.print", label: "Print", category: "File", bindings: [{ key: "p", mod: true }], overridesBrowser: true },
  // SPEC-066 AC-1/AC-2: the conventional Preferences accelerator, Cmd+, on
  // macOS / Ctrl+, elsewhere. Handled at the shell level so it opens from the
  // library and the editor alike.
  { id: "app.preferences", label: "Preferences", category: "File", bindings: [{ key: ",", mod: true }] },

  // View
  // SPEC-062 AC-1/AC-2: Ctrl/Cmd+K opens the command palette from anywhere in
  // the editor, handled at the document level alongside find/print. The earlier
  // Ctrl/Cmd+K = insert.link binding (SPEC-025) moves to Ctrl/Cmd+Alt+K to make
  // room, so the no-duplicate-binding invariant still holds. `k` is not a
  // RESERVED_BROWSER_KEY.
  { id: "app.commandPalette", label: "Search commands", category: "View", bindings: [{ key: "k", mod: true }] },
  { id: "view.shortcuts", label: "Keyboard shortcuts", category: "View", bindings: [{ key: "/", mod: true }] },
  {
    id: "view.formattingMarks",
    label: "Show formatting marks",
    category: "View",
    bindings: [{ key: "p", mod: true, shift: true }],
  },
  // SPEC-063 zoom (AC-11, AC-12). Ctrl/Cmd with =, - and 0; "+" (shift+=) is
  // an alternate for zoom in so both the labelled and the shifted key work.
  {
    id: "view.zoomIn",
    label: "Zoom in",
    category: "View",
    bindings: [
      { key: "=", mod: true },
      { key: "+", mod: true, shift: true },
    ],
  },
  { id: "view.zoomOut", label: "Zoom out", category: "View", bindings: [{ key: "-", mod: true }] },
  { id: "view.zoomReset", label: "Reset zoom to 100%", category: "View", bindings: [{ key: "0", mod: true }] },
  // SPEC-063 Focus mode (AC-46).
  { id: "view.focus", label: "Focus mode", category: "View", bindings: [{ key: "f", mod: true, shift: true }] },
];

function normalizeKey(key: string): string {
  return key.length === 1 ? key.toLowerCase() : key;
}

/** Whether the event is the binding's key: its physical `code` when it names
 * one (SPEC-123), otherwise the character or named key it produced. */
function keyMatches(binding: ShortcutBinding, event: KeyboardEvent): boolean {
  return binding.code ? event.code === binding.code : normalizeKey(binding.key) === normalizeKey(event.key);
}

/** Ctrl/Cmd as the binding asks. Without `ctrl` it is SPEC-021's rule: the
 * platform's own modifier is held exactly when `mod` is set, and the other
 * key is not looked at. With `ctrl` (SPEC-123) Control must be held, and on
 * macOS Cmd exactly when `mod` is also set; on Windows/Linux `mod` is Ctrl
 * itself, so `ctrl` satisfies it. */
function primaryModifiersMatch(binding: ShortcutBinding, event: KeyboardEvent, isMac: boolean): boolean {
  if (!binding.ctrl) {
    const modPressed = isMac ? event.metaKey : event.ctrlKey;
    return modPressed === !!binding.mod;
  }
  return event.ctrlKey && (!isMac || event.metaKey === !!binding.mod);
}

function bindingMatchesEvent(binding: ShortcutBinding, event: KeyboardEvent, isMac: boolean): boolean {
  if (!keyMatches(binding, event) || !primaryModifiersMatch(binding, event, isMac)) {
    return false;
  }
  if (event.altKey !== !!binding.alt) {
    return false;
  }
  return binding.shiftOptional || event.shiftKey === !!binding.shift;
}

/** Looks up the registry entry a keydown event triggers, if any — the only
 * way any part of the app should decide "was a shortcut just pressed"
 * (AC-2). Platform defaults to the runtime-resolved current platform; a
 * caller may override it (tests, or a future preview of the other
 * platform's bindings). `definitions` defaults to this registry; a product
 * passes its own table (SPEC-123's SHEET_SHORTCUTS) to match that instead. */
export function findShortcut(
  event: KeyboardEvent,
  isMac: boolean = isMacPlatform(),
  definitions: ShortcutDefinition[] = SHORTCUTS,
): ShortcutDefinition | undefined {
  return definitions.find((definition) => definition.bindings.some((binding) => bindingMatchesEvent(binding, event, isMac)));
}

function displayKey(key: string): string {
  return key.length === 1 ? key.toUpperCase() : key;
}

/** Renders one binding as display text, e.g. "Ctrl+Shift+X" / "Cmd+Shift+X".
 * `ctrl` leads, as "Ctrl" or (macOS) "Control" — "Ctrl+Shift+~" /
 * "Control+Shift+~" (SPEC-123); on Windows/Linux it is the same key as
 * `mod`'s Ctrl, so a binding with both names it once. */
export function formatBinding(binding: ShortcutBinding, isMac: boolean = isMacPlatform()): string {
  const parts: string[] = [];
  if (binding.ctrl) {
    parts.push(isMac ? "Control" : "Ctrl");
  }
  if (binding.mod && (isMac || !binding.ctrl)) {
    parts.push(isMac ? "Cmd" : "Ctrl");
  }
  if (binding.alt) {
    parts.push(isMac ? "Option" : "Alt");
  }
  if (binding.shift) {
    parts.push("Shift");
  }
  parts.push(displayKey(binding.key));
  return parts.join("+");
}

/** The display text for a command's primary binding — what toolbar tooltips
 * read (AC-15, AC-16). Returns "" for an unknown id rather than throwing, so
 * a stale id fails a visible test instead of crashing a render.
 * `definitions` as findShortcut's: this registry unless a product passes its
 * own table. */
export function shortcutLabel(id: string, isMac: boolean = isMacPlatform(), definitions: ShortcutDefinition[] = SHORTCUTS): string {
  const definition = definitions.find((entry) => entry.id === id);
  return definition ? formatBinding(definition.bindings[0], isMac) : "";
}

export interface ShortcutGroup {
  category: string;
  items: ShortcutDefinition[];
}

/** Groups definitions by category, preserving each category's first-seen
 * order in `definitions` — what the reference dialog renders (AC-7). */
export function groupShortcuts(definitions: ShortcutDefinition[] = SHORTCUTS): ShortcutGroup[] {
  const order: string[] = [];
  const byCategory = new Map<string, ShortcutDefinition[]>();
  for (const definition of definitions) {
    let items = byCategory.get(definition.category);
    if (!items) {
      items = [];
      byCategory.set(definition.category, items);
      order.push(definition.category);
    }
    items.push(definition);
  }
  return order.map((category) => ({ category, items: byCategory.get(category) ?? [] }));
}

/** Filters the registry for the dialog's search box (AC-6): matches the
 * command label, its category, or its bound key text. */
export function searchShortcuts(
  query: string,
  definitions: ShortcutDefinition[] = SHORTCUTS,
  isMac: boolean = isMacPlatform(),
): ShortcutDefinition[] {
  const trimmed = query.trim().toLowerCase();
  if (trimmed === "") {
    return definitions;
  }
  return definitions.filter((definition) => {
    const bindingText = definition.bindings.map((binding) => formatBinding(binding, isMac).toLowerCase()).join(" ");
    return (
      definition.label.toLowerCase().includes(trimmed) ||
      definition.category.toLowerCase().includes(trimmed) ||
      bindingText.includes(trimmed)
    );
  });
}

/** Whether a keydown's target lies outside the document editing surface —
 * inside an open dialog, the find bar, or any other text input (AC-12,
 * AC-13, AC-14) — meaning no registry shortcut should act on it. The main
 * contenteditable surface is exempted via `isContentEditable` (it's the only
 * contenteditable element in the app) so formatting/undo/etc keep working
 * while typing in the document itself. */
export function isOutsideDocumentInput(target: EventTarget | null): boolean {
  if (!(target instanceof Element)) {
    return false;
  }
  if (target.closest('[role="dialog"], [role="alertdialog"]') || target.closest('[data-testid="find-bar"]')) {
    return true;
  }
  if (target instanceof HTMLElement && target.isContentEditable) {
    return false;
  }
  return target.tagName === "INPUT" || target.tagName === "TEXTAREA" || target.tagName === "SELECT";
}
