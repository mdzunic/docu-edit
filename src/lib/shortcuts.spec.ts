import { describe, expect, it } from "vitest";
import {
  findShortcut,
  formatBinding,
  groupShortcuts,
  isOutsideDocumentInput,
  RESERVED_BROWSER_KEYS,
  searchShortcuts,
  SHORTCUTS,
  shortcutLabel,
} from "./shortcuts";

function keydown(init: Partial<KeyboardEventInit> & { key: string }): KeyboardEvent {
  return new KeyboardEvent("keydown", { bubbles: true, cancelable: true, ...init });
}

describe("SHORTCUTS registry (SPEC-021 AC-1, AC-5, AC-11)", () => {
  it("declares no two commands sharing the same binding", () => {
    const seen = new Map<string, string>();
    for (const definition of SHORTCUTS) {
      for (const binding of definition.bindings) {
        // shiftOptional matches both shift states, so it claims both
        // normalized slots — a future binding requiring/forbidding shift on
        // the same key would collide with it either way. Alt is part of the
        // normalization exactly as bindingMatchesEvent matches it (modifiers
        // are exact, never a subset): Ctrl+Alt+F (SPEC-031's insert
        // footnote) and Ctrl+F (find) are distinct bindings by design, the
        // same way the SPEC-019 style shortcuts Ctrl+Alt+1..6 always were.
        const shiftStates = binding.shiftOptional ? [false, true] : [!!binding.shift];
        for (const shiftHeld of shiftStates) {
          const normalized = `${binding.mod ? "mod+" : ""}${binding.alt ? "alt+" : ""}${shiftHeld ? "shift+" : ""}${binding.key.toLowerCase()}`;
          const owner = seen.get(normalized);
          expect(owner === undefined || owner === definition.id).toBe(true);
          seen.set(normalized, definition.id);
        }
      }
    }
  });

  it("never binds the reserved browser shortcuts Ctrl/Cmd+T/W/N", () => {
    for (const definition of SHORTCUTS) {
      for (const binding of definition.bindings) {
        if (binding.mod) {
          expect(RESERVED_BROWSER_KEYS).not.toContain(binding.key.toLowerCase());
        }
      }
    }
  });

  it("registers app.commandPalette under the View category (SPEC-062 AC-2, AC-3)", () => {
    const palette = SHORTCUTS.find((definition) => definition.id === "app.commandPalette");
    expect(palette).toBeDefined();
    expect(palette?.category).toBe("View");
    expect(palette?.label).toBe("Search commands");
    // It appears in the reference dialog because groupShortcuts renders every
    // registry entry; a real binding gives it a key there. SPEC-062 AC-1 puts
    // the palette on Ctrl/Cmd+K.
    expect(shortcutLabel("app.commandPalette", false)).toBe("Ctrl+K");
    expect(findShortcut(keydown({ key: "k", ctrlKey: true }), false)?.id).toBe("app.commandPalette");
    // SPEC-025's Insert link accelerator moved to Ctrl/Cmd+Alt+K to make room,
    // so both bindings coexist and neither collides.
    expect(shortcutLabel("insert.link", false)).toBe("Ctrl+Alt+K");
    expect(findShortcut(keydown({ key: "k", ctrlKey: true, altKey: true }), false)?.id).toBe("insert.link");
  });

  it("marks exactly the five deliberately-overridden browser shortcuts", () => {
    const overridden = SHORTCUTS.filter((definition) => definition.overridesBrowser).map((definition) => definition.id);
    expect(overridden.toSorted()).toEqual(
      ["file.print", "find.open", "format.bold", "format.italic", "format.underline"].toSorted(),
    );
  });
});

describe("findShortcut", () => {
  it("matches a plain Ctrl+B on a non-Mac platform", () => {
    const shortcut = findShortcut(keydown({ key: "b", ctrlKey: true }), false);
    expect(shortcut?.id).toBe("format.bold");
  });

  it("matches Cmd+B on a Mac platform, not Ctrl+B", () => {
    expect(findShortcut(keydown({ key: "b", metaKey: true }), true)?.id).toBe("format.bold");
    expect(findShortcut(keydown({ key: "b", ctrlKey: true }), true)).toBeUndefined();
  });

  it("does not match Ctrl+B on a Mac (wrong modifier for the platform)", () => {
    expect(findShortcut(keydown({ key: "b", ctrlKey: true }), true)).toBeUndefined();
  });

  it("distinguishes Ctrl+Z from Ctrl+Shift+Z", () => {
    expect(findShortcut(keydown({ key: "z", ctrlKey: true }), false)?.id).toBe("edit.undo");
    expect(findShortcut(keydown({ key: "z", ctrlKey: true, shiftKey: true }), false)?.id).toBe("edit.redo");
  });

  it("matches redo's alternate Ctrl+Y binding", () => {
    expect(findShortcut(keydown({ key: "y", ctrlKey: true }), false)?.id).toBe("edit.redo");
  });

  it("matches named keys like Tab and PageDown", () => {
    expect(findShortcut(keydown({ key: "Tab" }), false)?.id).toBe("paragraph.indent");
    expect(findShortcut(keydown({ key: "Tab", shiftKey: true }), false)?.id).toBe("paragraph.outdent");
    expect(findShortcut(keydown({ key: "PageDown" }), false)?.id).toBe("nav.pageDown");
  });

  it("returns undefined for an unbound combination", () => {
    // Ctrl+K was this test's example until SPEC-025 bound it to "Insert
    // link"; Ctrl+Q remains genuinely unbound.
    expect(findShortcut(keydown({ key: "q", ctrlKey: true }), false)).toBeUndefined();
  });

  it("matches Ctrl+Home and Ctrl+Shift+Home as the same command (shift extends, doesn't gate)", () => {
    expect(findShortcut(keydown({ key: "Home", ctrlKey: true }), false)?.id).toBe("nav.documentStart");
    expect(findShortcut(keydown({ key: "Home", ctrlKey: true, shiftKey: true }), false)?.id).toBe("nav.documentStart");
  });

  it("matches PageDown with or without Shift as the same command", () => {
    expect(findShortcut(keydown({ key: "PageDown" }), false)?.id).toBe("nav.pageDown");
    expect(findShortcut(keydown({ key: "PageDown", shiftKey: true }), false)?.id).toBe("nav.pageDown");
  });
});

describe("formatBinding / shortcutLabel", () => {
  it("renders Ctrl on non-Mac and Cmd on Mac", () => {
    expect(formatBinding({ key: "b", mod: true }, false)).toBe("Ctrl+B");
    expect(formatBinding({ key: "b", mod: true }, true)).toBe("Cmd+B");
  });

  it("includes Shift when the binding requires it", () => {
    expect(formatBinding({ key: "x", mod: true, shift: true }, false)).toBe("Ctrl+Shift+X");
  });

  it("leaves named keys as-is", () => {
    expect(formatBinding({ key: "PageDown" }, false)).toBe("PageDown");
  });

  it("shortcutLabel reads the first (primary) binding", () => {
    expect(shortcutLabel("edit.redo", false)).toBe("Ctrl+Shift+Z");
  });

  it("shortcutLabel returns an empty string for an unknown id", () => {
    expect(shortcutLabel("no.such.command", false)).toBe("");
  });
});

describe("groupShortcuts", () => {
  it("preserves each category's first-seen order and keeps every entry", () => {
    const groups = groupShortcuts();
    const total = groups.reduce((sum, group) => sum + group.items.length, 0);
    expect(total).toBe(SHORTCUTS.length);
    const categories = groups.map((group) => group.category);
    expect(new Set(categories).size).toBe(categories.length);
  });
});

describe("searchShortcuts", () => {
  it("returns everything for a blank query", () => {
    expect(searchShortcuts("")).toHaveLength(SHORTCUTS.length);
  });

  it("matches by label", () => {
    const results = searchShortcuts("bold");
    expect(results.map((r) => r.id)).toEqual(["format.bold"]);
  });

  it("matches by category", () => {
    const results = searchShortcuts("navigation", SHORTCUTS, false);
    expect(results.every((r) => r.category === "Navigation")).toBe(true);
    expect(results.length).toBeGreaterThan(0);
  });

  it("matches by bound key text", () => {
    const results = searchShortcuts("ctrl+shift+z", SHORTCUTS, false);
    expect(results.map((r) => r.id)).toContain("edit.redo");
  });
});

describe("isOutsideDocumentInput (AC-12, AC-13, AC-14)", () => {
  it("is false for a null/non-element target", () => {
    expect(isOutsideDocumentInput(null)).toBe(false);
  });

  it("is false for the document's contenteditable surface", () => {
    const el = document.createElement("div");
    el.contentEditable = "true";
    document.body.append(el);
    expect(isOutsideDocumentInput(el)).toBe(false);
    el.remove();
  });

  it("is true inside an open dialog", () => {
    const dialog = document.createElement("div");
    dialog.setAttribute("role", "dialog");
    const input = document.createElement("input");
    dialog.append(input);
    document.body.append(dialog);
    expect(isOutsideDocumentInput(input)).toBe(true);
    expect(isOutsideDocumentInput(dialog)).toBe(true);
    dialog.remove();
  });

  it("is true inside the find bar", () => {
    const bar = document.createElement("div");
    bar.setAttribute("data-testid", "find-bar");
    const input = document.createElement("input");
    bar.append(input);
    document.body.append(bar);
    expect(isOutsideDocumentInput(input)).toBe(true);
    bar.remove();
  });

  it("is true for a plain text input outside the document", () => {
    const input = document.createElement("input");
    document.body.append(input);
    expect(isOutsideDocumentInput(input)).toBe(true);
    input.remove();
  });

  it("is false for a non-input element outside any dialog or find bar", () => {
    const div = document.createElement("div");
    document.body.append(div);
    expect(isOutsideDocumentInput(div)).toBe(false);
    div.remove();
  });
});
