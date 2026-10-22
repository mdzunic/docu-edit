import { describe, expect, it } from "vitest";
import { SHEET_SHORTCUTS, sheetShortcutDefinitions, type SheetShortcutId } from "../views/spreadsheet/sheetShortcuts";
import { isMacPlatform } from "./core/platform";
import {
  findShortcut,
  formatBinding,
  groupShortcuts,
  isOutsideDocumentInput,
  RESERVED_BROWSER_KEYS,
  searchShortcuts,
  SHORTCUTS,
  shortcutLabel,
  type ShortcutBinding,
  type ShortcutDefinition,
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

// SPEC-123: a product's own table (the sheet's) through the same matcher, the
// `code`/`ctrl` binding fields, and the text registry exactly as it was.

const PLATFORMS = [
  { isMac: false, name: "Windows/Linux" },
  { isMac: true, name: "macOS" },
];

const SHEET_IDS: SheetShortcutId[] = [
  "sheet.formatCells",
  "sheet.bold",
  "sheet.italic",
  "sheet.underline",
  "sheet.strikethrough",
  "sheet.formatGeneral",
  "sheet.formatNumber",
  "sheet.formatTime",
  "sheet.formatDate",
  "sheet.formatCurrency",
  "sheet.formatPercent",
  "sheet.formatScientific",
  "sheet.insertCells",
  "sheet.deleteCells",
  "sheet.pasteSpecial",
  "sheet.fillDown",
  "sheet.fillRight",
  "sheet.insertFunction",
  "sheet.toggleFilter",
  "sheet.reapplyFilter",
  "sheet.nameManager",
  "sheet.goTo",
  "sheet.note",
  "sheet.newComment",
  "sheet.insertLink",
];

/** The keydown a binding describes on one platform, for sweeping a whole
 * table: `mod` presses Ctrl or Cmd, `ctrl` presses Control, and a `code`
 * binding presses its physical key typing the character its label shows (a
 * US layout). */
function pressFor(binding: ShortcutBinding, isMac: boolean, shiftKey = !!binding.shift): KeyboardEvent {
  return keydown({
    key: binding.key,
    code: binding.code,
    ctrlKey: !!binding.ctrl || (!isMac && !!binding.mod),
    metaKey: isMac && !!binding.mod,
    altKey: !!binding.alt,
    shiftKey,
  });
}

/** The key presses a binding claims on one platform, normalised the way
 * bindingMatchesEvent matches: `mod` becomes the key it means there and
 * `ctrl` Control, one key on Windows/Linux. A binding without `ctrl` never
 * looks at Control on macOS, so there it claims both states of Control, as
 * shiftOptional claims both states of Shift. A `code` binding claims its
 * physical key and the character its label shows, so it also collides with a
 * key binding for that character. */
function claimedPresses(binding: ShortcutBinding, isMac: boolean): string[] {
  const controlStates = isMac ? (binding.ctrl ? [true] : [false, true]) : [!!(binding.ctrl || binding.mod)];
  const shiftStates = binding.shiftOptional ? [false, true] : [!!binding.shift];
  const targets = [`key:${binding.key.toLowerCase()}`, ...(binding.code ? [`code:${binding.code}`] : [])];
  const presses: string[] = [];
  for (const controlHeld of controlStates) {
    for (const shiftHeld of shiftStates) {
      const modifiers = `${controlHeld ? "ctrl+" : ""}${isMac && binding.mod ? "cmd+" : ""}${binding.alt ? "alt+" : ""}${shiftHeld ? "shift+" : ""}`;
      presses.push(...targets.map((target) => modifiers + target));
    }
  }
  return presses;
}

describe('SHEET_SHORTCUTS (SPEC-123 "Shortcuts")', () => {
  it("declares the sheet's twenty-five commands — SPEC-123's twelve under Sheet formatting, SPEC-124's insert/delete under Rows and columns, SPEC-125's paste special and fills under Paste and fill, SPEC-127's Insert function and SPEC-133's Name Manager under Formulas, SPEC-128's filter toggle and reapply under Sort and filter, SPEC-134's Go to under Navigation, SPEC-135's note and New comment under Notes and comments and its Insert link under Links — one binding each but paste special's and Go to's two", () => {
    expect(SHEET_SHORTCUTS).toEqual(sheetShortcutDefinitions(isMacPlatform()));
    const pasteAndFill = ["sheet.pasteSpecial", "sheet.fillDown", "sheet.fillRight"];
    const sortAndFilter = ["sheet.toggleFilter", "sheet.reapplyFilter"];
    const annotations = ["sheet.note", "sheet.newComment"];
    for (const { isMac } of PLATFORMS) {
      const definitions = sheetShortcutDefinitions(isMac);
      expect(definitions.map((definition) => definition.id)).toEqual(SHEET_IDS);
      for (const definition of definitions) {
        const structural = definition.id === "sheet.insertCells" || definition.id === "sheet.deleteCells";
        const category =
          (definition.id === "sheet.insertFunction" || definition.id === "sheet.nameManager")
            ? "Formulas"
            : annotations.includes(definition.id)
              ? "Notes and comments"
            : definition.id === "sheet.insertLink"
              ? "Links"
            : pasteAndFill.includes(definition.id)
              ? "Paste and fill"
              : sortAndFilter.includes(definition.id)
                ? "Sort and filter"
                : definition.id === "sheet.goTo"
                  ? "Navigation"
                  : "Sheet formatting";
        expect(definition.category).toBe(structural ? "Rows and columns" : category);
        expect(definition.bindings).toHaveLength(definition.id === "sheet.pasteSpecial" || definition.id === "sheet.goTo" ? 2 : 1);
      }
    }
  });

  it("binds SPEC-125's paste special to Ctrl/Cmd+Shift+V and Ctrl+Alt+V (Control+Option+V on macOS) by physical key, fills to Ctrl/Cmd+D and Ctrl/Cmd+R", () => {
    for (const { isMac } of PLATFORMS) {
      const definitions = sheetShortcutDefinitions(isMac);
      const mod = isMac ? { metaKey: true } : { ctrlKey: true };
      expect(findShortcut(keydown({ key: "V", code: "KeyV", shiftKey: true, ...mod }), isMac, definitions)?.id).toBe("sheet.pasteSpecial");
      // Option+V types "√" on macOS: the key matches by code.
      expect(findShortcut(keydown({ key: isMac ? "√" : "v", code: "KeyV", ctrlKey: true, altKey: true }), isMac, definitions)?.id).toBe("sheet.pasteSpecial");
      expect(findShortcut(keydown({ key: "v", code: "KeyV", ...mod }), isMac, definitions)).toBeUndefined();
      expect(findShortcut(keydown({ key: "d", code: "KeyD", ...mod }), isMac, definitions)?.id).toBe("sheet.fillDown");
      expect(findShortcut(keydown({ key: "r", code: "KeyR", ...mod }), isMac, definitions)?.id).toBe("sheet.fillRight");
    }
    // Cmd+Option+V on macOS is no binding of the sheet's.
    expect(findShortcut(keydown({ key: "√", code: "KeyV", metaKey: true, altKey: true }), true, sheetShortcutDefinitions(true))).toBeUndefined();
  });

  it("binds SPEC-128's Filter to Ctrl/Cmd+Shift+L and Reapply filter to Ctrl/Cmd+Alt+L by physical key", () => {
    for (const { isMac } of PLATFORMS) {
      const definitions = sheetShortcutDefinitions(isMac);
      const mod = isMac ? { metaKey: true } : { ctrlKey: true };
      expect(findShortcut(keydown({ key: "L", code: "KeyL", shiftKey: true, ...mod }), isMac, definitions)?.id).toBe("sheet.toggleFilter");
      // Option+L types "¬" on macOS: the key matches by code.
      expect(findShortcut(keydown({ key: isMac ? "¬" : "l", code: "KeyL", altKey: true, ...mod }), isMac, definitions)?.id).toBe("sheet.reapplyFilter");
      expect(findShortcut(keydown({ key: "l", code: "KeyL", ...mod }), isMac, definitions)).toBeUndefined();
    }
  });

  it("binds SPEC-134's Go to to Control+G on every platform and to F5 — macOS keeps Cmd+G", () => {
    for (const { isMac } of PLATFORMS) {
      const definitions = sheetShortcutDefinitions(isMac);
      expect(findShortcut(keydown({ key: "g", code: "KeyG", ctrlKey: true }), isMac, definitions)?.id).toBe("sheet.goTo");
      expect(findShortcut(keydown({ key: "F5", code: "F5" }), isMac, definitions)?.id).toBe("sheet.goTo");
      expect(findShortcut(keydown({ key: "g", code: "KeyG", ctrlKey: true, shiftKey: true }), isMac, definitions)).toBeUndefined();
    }
    // Cmd+G (find-next on macOS) and Cmd+Control+G are no binding of the sheet's.
    expect(findShortcut(keydown({ key: "g", code: "KeyG", metaKey: true }), true, sheetShortcutDefinitions(true))).toBeUndefined();
    expect(findShortcut(keydown({ key: "g", code: "KeyG", metaKey: true, ctrlKey: true }), true, sheetShortcutDefinitions(true))).toBeUndefined();
  });

  it("binds SPEC-135's note to Shift+F2 and New comment to Ctrl+Alt+M / Cmd+Option+M by physical key", () => {
    for (const { isMac } of PLATFORMS) {
      const definitions = sheetShortcutDefinitions(isMac);
      const mod = isMac ? { metaKey: true } : { ctrlKey: true };
      expect(findShortcut(keydown({ key: "F2", code: "F2", shiftKey: true }), isMac, definitions)?.id).toBe("sheet.note");
      expect(findShortcut(keydown({ key: "F2", code: "F2" }), isMac, definitions)).toBeUndefined();
      // Option+M types "µ" on macOS: the key matches by code.
      expect(findShortcut(keydown({ key: isMac ? "µ" : "m", code: "KeyM", altKey: true, ...mod }), isMac, definitions)?.id).toBe("sheet.newComment");
      expect(findShortcut(keydown({ key: "m", code: "KeyM", ...mod }), isMac, definitions)).toBeUndefined();
    }
  });

  it("binds SPEC-135's Insert link to Ctrl/Cmd+K, and not Ctrl/Cmd+Alt+K or Ctrl/Cmd+Shift+K", () => {
    for (const { isMac } of PLATFORMS) {
      const definitions = sheetShortcutDefinitions(isMac);
      const mod = isMac ? { metaKey: true } : { ctrlKey: true };
      expect(findShortcut(keydown({ key: "k", code: "KeyK", ...mod }), isMac, definitions)?.id).toBe("sheet.insertLink");
      expect(findShortcut(keydown({ key: "k", code: "KeyK", altKey: true, ...mod }), isMac, definitions)).toBeUndefined();
      expect(findShortcut(keydown({ key: "K", code: "KeyK", shiftKey: true, ...mod }), isMac, definitions)).toBeUndefined();
    }
  });

  it("binds SPEC-127's Insert function to Shift+F3 on both platforms", () => {
    for (const { isMac } of PLATFORMS) {
      const definitions = sheetShortcutDefinitions(isMac);
      expect(findShortcut(keydown({ key: "F3", code: "F3", shiftKey: true }), isMac, definitions)?.id).toBe("sheet.insertFunction");
      expect(findShortcut(keydown({ key: "F3", code: "F3" }), isMac, definitions)).toBeUndefined();
    }
  });

  it("binds SPEC-124's insert and delete to Ctrl+Alt+= / Ctrl+Alt+- (Cmd+Option on macOS) by physical key, apart from the zoom keys", () => {
    for (const { isMac } of PLATFORMS) {
      const definitions = sheetShortcutDefinitions(isMac);
      const mod = isMac ? { metaKey: true } : { ctrlKey: true };
      // Option's dead characters on macOS: the key still matches by code.
      expect(findShortcut(keydown({ key: isMac ? "≠" : "=", code: "Equal", altKey: true, ...mod }), isMac, definitions)?.id).toBe("sheet.insertCells");
      expect(findShortcut(keydown({ key: isMac ? "–" : "-", code: "Minus", altKey: true, ...mod }), isMac, definitions)?.id).toBe("sheet.deleteCells");
      // Ctrl/Cmd+= and Ctrl/Cmd+- without Alt stay the zoom.
      expect(findShortcut(keydown({ key: "=", code: "Equal", ...mod }), isMac, definitions)).toBeUndefined();
      expect(findShortcut(keydown({ key: "=", code: "Equal", ...mod }), isMac)?.id).toBe("view.zoomIn");
      expect(findShortcut(keydown({ key: "-", code: "Minus", ...mod }), isMac)?.id).toBe("view.zoomOut");
    }
  });

  it("never binds the reserved browser shortcuts Ctrl/Cmd+T/W/N, on either platform", () => {
    const reservedCodes = RESERVED_BROWSER_KEYS.map((key) => `Key${key.toUpperCase()}`);
    for (const definitions of [SHEET_SHORTCUTS, sheetShortcutDefinitions(false), sheetShortcutDefinitions(true)]) {
      for (const definition of definitions) {
        for (const binding of definition.bindings) {
          if (binding.mod || binding.ctrl) {
            expect(RESERVED_BROWSER_KEYS).not.toContain(binding.key.toLowerCase());
            expect(reservedCodes).not.toContain(binding.code);
          }
        }
      }
    }
  });

  it("declares no two commands sharing a key press, on either platform", () => {
    for (const { isMac, name } of PLATFORMS) {
      const seen = new Map<string, string>();
      for (const definition of sheetShortcutDefinitions(isMac)) {
        for (const binding of definition.bindings) {
          for (const press of claimedPresses(binding, isMac)) {
            const owner = seen.get(press);
            expect(owner === undefined || owner === definition.id, `${name}: ${press} is ${owner} and ${definition.id}`).toBe(true);
            seen.set(press, definition.id);
          }
        }
      }
    }
  });

  it("matches every binding to its own command on its own platform", () => {
    for (const { isMac, name } of PLATFORMS) {
      const definitions = sheetShortcutDefinitions(isMac);
      for (const definition of definitions) {
        for (const binding of definition.bindings) {
          expect(findShortcut(pressFor(binding, isMac), isMac, definitions)?.id, `${name}: ${definition.id}`).toBe(definition.id);
        }
      }
    }
  });
});

describe("findShortcut with code and ctrl bindings (SPEC-123)", () => {
  const windows = sheetShortcutDefinitions(false);
  const mac = sheetShortcutDefinitions(true);

  it("matches Ctrl+Shift+1 on Windows/Linux and Control+Shift+1 on macOS to the number format", () => {
    const press = keydown({ key: "!", code: "Digit1", ctrlKey: true, shiftKey: true });
    expect(findShortcut(press, false, windows)?.id).toBe("sheet.formatNumber");
    expect(findShortcut(press, true, mac)?.id).toBe("sheet.formatNumber");
  });

  it("never takes Cmd for Control on macOS: Cmd+Shift+1 and Control+Cmd+Shift+1 match nothing", () => {
    // Either character a browser may report for Cmd+Shift+1.
    expect(findShortcut(keydown({ key: "!", code: "Digit1", metaKey: true, shiftKey: true }), true, mac)).toBeUndefined();
    expect(findShortcut(keydown({ key: "1", code: "Digit1", metaKey: true, shiftKey: true }), true, mac)).toBeUndefined();
    expect(findShortcut(keydown({ key: "!", code: "Digit1", ctrlKey: true, metaKey: true, shiftKey: true }), true, mac)).toBeUndefined();
  });

  it("matches the number formats by physical key, whatever character the layout types there", () => {
    // German QWERTZ: Shift+1 is still "!", but Shift+2 is '"', Shift+3 "§",
    // Shift+6 "&", and the key left of 1 (Backquote) types "°".
    const german = [
      { key: "!", code: "Digit1", id: "sheet.formatNumber" },
      { key: '"', code: "Digit2", id: "sheet.formatTime" },
      { key: "§", code: "Digit3", id: "sheet.formatDate" },
      { key: "&", code: "Digit6", id: "sheet.formatScientific" },
      { key: "°", code: "Backquote", id: "sheet.formatGeneral" },
    ];
    for (const { key, code, id } of german) {
      const press = keydown({ key, code, ctrlKey: true, shiftKey: true });
      expect(findShortcut(press, false, windows)?.id, `${code} typing ${key}`).toBe(id);
      expect(findShortcut(press, true, mac)?.id, `${code} typing ${key}`).toBe(id);
    }
    // ...and a "!" typed by some other key is not the 1 key.
    expect(findShortcut(keydown({ key: "!", code: "Slash", ctrlKey: true, shiftKey: true }), false, windows)).toBeUndefined();
  });

  it("keeps Alt and Shift exact for ctrl bindings", () => {
    // AltGr arrives as Ctrl+Alt on Windows; it never triggers a number format.
    expect(findShortcut(keydown({ key: "!", code: "Digit1", ctrlKey: true, altKey: true, shiftKey: true }), false, windows)).toBeUndefined();
    // Without Shift, Control+1 on macOS is nothing (Format cells is Cmd+1).
    expect(findShortcut(keydown({ key: "1", code: "Digit1", ctrlKey: true }), true, mac)).toBeUndefined();
  });

  it("reads ctrl with mod as Control+Cmd on macOS and as plain Ctrl on Windows/Linux", () => {
    const table: ShortcutDefinition[] = [{ id: "test.both", label: "Both", category: "Test", bindings: [{ key: "k", ctrl: true, mod: true }] }];
    expect(findShortcut(keydown({ key: "k", ctrlKey: true }), false, table)?.id).toBe("test.both");
    expect(findShortcut(keydown({ key: "k", ctrlKey: true, metaKey: true }), true, table)?.id).toBe("test.both");
    expect(findShortcut(keydown({ key: "k", ctrlKey: true }), true, table)).toBeUndefined();
    expect(findShortcut(keydown({ key: "k", metaKey: true }), true, table)).toBeUndefined();
  });

  it("matches Format cells on Ctrl+1 (Windows/Linux) and Cmd+1 (macOS)", () => {
    expect(findShortcut(keydown({ key: "1", code: "Digit1", ctrlKey: true }), false, windows)?.id).toBe("sheet.formatCells");
    expect(findShortcut(keydown({ key: "1", code: "Digit1", metaKey: true }), true, mac)?.id).toBe("sheet.formatCells");
  });

  it("tells Windows/Linux Ctrl+5 (strikethrough) from Ctrl+Shift+5 (percent)", () => {
    expect(findShortcut(keydown({ key: "5", code: "Digit5", ctrlKey: true }), false, windows)?.id).toBe("sheet.strikethrough");
    expect(findShortcut(keydown({ key: "%", code: "Digit5", ctrlKey: true, shiftKey: true }), false, windows)?.id).toBe("sheet.formatPercent");
    // The text editor's Ctrl+Shift+X is not the sheet's strikethrough here.
    expect(findShortcut(keydown({ key: "X", code: "KeyX", ctrlKey: true, shiftKey: true }), false, windows)).toBeUndefined();
  });

  it("matches macOS strikethrough on Cmd+Shift+X, not Cmd+5", () => {
    expect(findShortcut(keydown({ key: "x", code: "KeyX", metaKey: true, shiftKey: true }), true, mac)?.id).toBe("sheet.strikethrough");
    expect(findShortcut(keydown({ key: "5", code: "Digit5", metaKey: true }), true, mac)).toBeUndefined();
  });

  it("matches only the definitions it is given", () => {
    const ctrlB = keydown({ key: "b", code: "KeyB", ctrlKey: true });
    expect(findShortcut(ctrlB, false, windows)?.id).toBe("sheet.bold");
    expect(findShortcut(ctrlB, false)?.id).toBe("format.bold");
    expect(findShortcut(ctrlB, false, [])).toBeUndefined();
  });
});

describe("the text registry after SPEC-123", () => {
  it("still matches every binding to its own command on both platforms, by default and when passed explicitly", () => {
    for (const { isMac, name } of PLATFORMS) {
      for (const definition of SHORTCUTS) {
        for (const binding of definition.bindings) {
          for (const shiftHeld of binding.shiftOptional ? [false, true] : [!!binding.shift]) {
            const press = pressFor(binding, isMac, shiftHeld);
            expect(findShortcut(press, isMac)?.id, `${name}: ${definition.id}`).toBe(definition.id);
            expect(findShortcut(press, isMac, SHORTCUTS)?.id, `${name}: ${definition.id}`).toBe(definition.id);
          }
        }
      }
    }
  });

  it("never answers with a sheet command: the sheet's presses are the text editor's own commands on the same keys, or nothing", () => {
    expect(SHORTCUTS.some((definition) => definition.id.startsWith("sheet."))).toBe(false);
    // Ctrl/Cmd+B/I/U, and macOS Cmd+Shift+X, are the text editor's own
    // commands on the same keys — as SPEC-125's Ctrl/Cmd+Shift+V (paste
    // without formatting) and Ctrl/Cmd+R (align right) are, and SPEC-135's
    // Ctrl+Alt+M / Cmd+Option+M (insert comment, the same command), and its
    // Ctrl/Cmd+K is the text side's command palette (a view the workbook
    // never mounts it in); no other sheet press means anything to it.
    const sameKeys = new Map([
      ["sheet.bold", "format.bold"],
      ["sheet.italic", "format.italic"],
      ["sheet.underline", "format.underline"],
      ["sheet.pasteSpecial", "edit.pastePlain"],
      ["sheet.fillRight", "paragraph.alignRight"],
      ["sheet.newComment", "insert.comment"],
      ["sheet.insertLink", "app.commandPalette"],
    ]);
    for (const { isMac, name } of PLATFORMS) {
      for (const definition of sheetShortcutDefinitions(isMac)) {
        const expected = isMac && definition.id === "sheet.strikethrough" ? "format.strikethrough" : sameKeys.get(definition.id);
        expect(findShortcut(pressFor(definition.bindings[0], isMac), isMac)?.id, `${name}: ${definition.id}`).toBe(expected);
      }
    }
  });

  it("does not match Ctrl+Shift+1 (the 1 key) on Windows/Linux or Control+Shift+1 on macOS", () => {
    const press = keydown({ key: "!", code: "Digit1", ctrlKey: true, shiftKey: true });
    expect(findShortcut(press, false)).toBeUndefined();
    expect(findShortcut(press, true)).toBeUndefined();
  });

  it("still matches its bindings by the character typed, never the physical key", () => {
    // Dvorak types "b" on the physical N key, and "x" on the physical B key.
    expect(findShortcut(keydown({ key: "b", code: "KeyN", ctrlKey: true }), false)?.id).toBe("format.bold");
    expect(findShortcut(keydown({ key: "x", code: "KeyB", ctrlKey: true }), false)).toBeUndefined();
  });

  it("still looks only at the platform's own modifier for a binding without ctrl, as SPEC-021 matched it", () => {
    expect(findShortcut(keydown({ key: "b", ctrlKey: true, metaKey: true }), true)?.id).toBe("format.bold");
    expect(findShortcut(keydown({ key: "b", ctrlKey: true, metaKey: true }), false)?.id).toBe("format.bold");
  });
});

describe("formatBinding / shortcutLabel with ctrl (SPEC-123)", () => {
  it("labels the number formats Ctrl+Shift+<key> on Windows/Linux and Control+Shift+<key> on macOS", () => {
    expect(shortcutLabel("sheet.formatNumber", false, sheetShortcutDefinitions(false))).toBe("Ctrl+Shift+!");
    expect(shortcutLabel("sheet.formatNumber", true, sheetShortcutDefinitions(true))).toBe("Control+Shift+!");
    expect(shortcutLabel("sheet.formatGeneral", false, sheetShortcutDefinitions(false))).toBe("Ctrl+Shift+~");
    expect(shortcutLabel("sheet.formatGeneral", true, sheetShortcutDefinitions(true))).toBe("Control+Shift+~");
  });

  it("labels each platform's own strikethrough and Format cells", () => {
    expect(shortcutLabel("sheet.strikethrough", false, sheetShortcutDefinitions(false))).toBe("Ctrl+5");
    expect(shortcutLabel("sheet.strikethrough", true, sheetShortcutDefinitions(true))).toBe("Cmd+Shift+X");
    expect(shortcutLabel("sheet.formatCells", false, sheetShortcutDefinitions(false))).toBe("Ctrl+1");
    expect(shortcutLabel("sheet.formatCells", true, sheetShortcutDefinitions(true))).toBe("Cmd+1");
  });

  it("puts Control first, and names Ctrl once when a binding has both ctrl and mod", () => {
    expect(formatBinding({ key: "a", ctrl: true, alt: true, shift: true }, false)).toBe("Ctrl+Alt+Shift+A");
    expect(formatBinding({ key: "a", ctrl: true, alt: true, shift: true }, true)).toBe("Control+Option+Shift+A");
    expect(formatBinding({ key: "k", ctrl: true, mod: true }, false)).toBe("Ctrl+K");
    expect(formatBinding({ key: "k", ctrl: true, mod: true }, true)).toBe("Control+Cmd+K");
  });

  it("labels from the definitions it is given; the text registry has no sheet ids", () => {
    expect(shortcutLabel("sheet.bold", false)).toBe("");
    expect(shortcutLabel("sheet.bold", false, sheetShortcutDefinitions(false))).toBe("Ctrl+B");
    expect(shortcutLabel("format.bold", false, sheetShortcutDefinitions(false))).toBe("");
  });
});
