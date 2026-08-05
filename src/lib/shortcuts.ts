// The single keyboard-shortcut registry (SPEC-021). Every key binding the
// app intercepts is declared here — command id, display label, category and
// platform-neutral binding(s) — so the editor's keydown handlers, toolbar
// tooltips and the "Keyboard shortcuts" reference dialog all read the same
// source of truth (AC-1, AC-2). `mod` means "Ctrl on Windows/Linux, Cmd on
// macOS", resolved at runtime via `isMacPlatform()` — never by sniffing the
// user-agent string (AC-3, AC-4).
//
// Deliberately excluded: per-component dismissal conventions (Escape closes
// a menu/popover/dialog, Enter confirms a text field) aren't registered
// here. Those are a generic, universal UI convention every overlay in the
// app implements locally, not a named, documented command — real word
// processors don't list "Escape closes this dialog" in their shortcut
// reference either. The registry is for commands a user would look up by
// name.
import { isMacPlatform } from "./platform";

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
  mod?: boolean;
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
 * the app (AC-11) — asserted against this registry in shortcuts.spec.ts. */
export const RESERVED_BROWSER_KEYS = ["t", "w", "n"];

export const SHORTCUTS: ShortcutDefinition[] = [
  // Formatting
  { id: "format.bold", label: "Bold", category: "Formatting", bindings: [{ key: "b", mod: true }], overridesBrowser: true },
  { id: "format.italic", label: "Italic", category: "Formatting", bindings: [{ key: "i", mod: true }], overridesBrowser: true },
  { id: "format.underline", label: "Underline", category: "Formatting", bindings: [{ key: "u", mod: true }], overridesBrowser: true },
  { id: "format.strikethrough", label: "Strikethrough", category: "Formatting", bindings: [{ key: "x", mod: true, shift: true }] },
  { id: "format.subscript", label: "Subscript", category: "Formatting", bindings: [{ key: ",", mod: true }] },
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

  // File
  { id: "file.print", label: "Print", category: "File", bindings: [{ key: "p", mod: true }], overridesBrowser: true },

  // View
  { id: "view.shortcuts", label: "Keyboard shortcuts", category: "View", bindings: [{ key: "/", mod: true }] },
];

function normalizeKey(key: string): string {
  return key.length === 1 ? key.toLowerCase() : key;
}

function bindingMatchesEvent(binding: ShortcutBinding, event: KeyboardEvent, isMac: boolean): boolean {
  const modPressed = isMac ? event.metaKey : event.ctrlKey;
  if (normalizeKey(binding.key) !== normalizeKey(event.key) || modPressed !== !!binding.mod) {
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
 * platform's bindings). */
export function findShortcut(event: KeyboardEvent, isMac: boolean = isMacPlatform()): ShortcutDefinition | undefined {
  return SHORTCUTS.find((definition) => definition.bindings.some((binding) => bindingMatchesEvent(binding, event, isMac)));
}

function displayKey(key: string): string {
  return key.length === 1 ? key.toUpperCase() : key;
}

/** Renders one binding as display text, e.g. "Ctrl+Shift+X" / "Cmd+Shift+X". */
export function formatBinding(binding: ShortcutBinding, isMac: boolean = isMacPlatform()): string {
  const parts: string[] = [];
  if (binding.mod) {
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
 * a stale id fails a visible test instead of crashing a render. */
export function shortcutLabel(id: string, isMac: boolean = isMacPlatform()): string {
  const definition = SHORTCUTS.find((entry) => entry.id === id);
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
  if (target.closest('[role="dialog"]') || target.closest('[data-testid="find-bar"]')) {
    return true;
  }
  if (target instanceof HTMLElement && target.isContentEditable) {
    return false;
  }
  return target.tagName === "INPUT" || target.tagName === "TEXTAREA" || target.tagName === "SELECT";
}
