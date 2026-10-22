---
id: SPEC-049
title: Menu bar and toolbar consolidation
status: draft
area: shell
depends_on: [SPEC-001, SPEC-021, SPEC-037]
---
## Why
SPEC-001 built one toolbar strip so "every later feature adds a control to a
known place instead of inventing its own layout". Thirty-seven specs later that
place holds forty-three top-level controls and scrolls sideways past the edge of
a 1280px window, so the known place is no longer findable. Seven dropdown menus
have already been invented ad hoc — Insert break, Notes, Header/footer,
Multilevel list, Autocorrect, Download, and the one View menu SPEC-037 actually
asked for — each by a spec that needed somewhere to put something and had
nowhere to put it.

This spec ratifies that drift into a menu bar, and makes it the answer to
"where does my control go?" for every spec after it. That question is live
right now: SPEC-039 adds another toolbar toggle, and SPEC-043 already writes
"Tools → Check accessibility" against a Tools menu nothing creates.

The toolbar does not go away. It keeps the controls people use continuously —
the ones worth permanent screen space — and stops being the dumping ground for
the ones they use twice a document.

## Acceptance criteria
- [ ] A menu bar sits between the app bar and the toolbar strip, with the menus File, Edit, View, Insert, Format and Tools, in that order
- [ ] Every command currently reachable from the toolbar strip is reachable from exactly one menu, and no command appears in two menus
- [ ] The toolbar strip retains only the continuous-use controls — undo/redo, the style picker, font family and size, bold/italic/underline/strikethrough, text colour and highlight, and the paragraph alignment/list controls — and no longer overflows at a 1280px viewport
- [ ] The seven existing ad-hoc dropdowns (Insert break, Notes, Header/footer, Multilevel list, Autocorrect, View, and the table insert control) are absorbed as sections of the menus above, and disappear as separate toolbar buttons
- [ ] A menu item for a command that has a keyboard shortcut shows that shortcut's label, read from the SPEC-021 registry rather than written out again in the menu
- [ ] A command that does not apply to the current selection is present but disabled, never hidden — the menus are a stable map of the product, not a changing one
- [ ] The menu bar is one tab stop; Left/Right move between menus, Up/Down move within an open menu, Enter or Space activates, Escape closes and returns focus to the menu title
- [ ] Opening one menu and moving the pointer across the bar opens the neighbouring menu without a second click, the standard menu-bar behaviour
- [ ] The menu bar is exposed as `role="menubar"` with `role="menu"` popups and `role="menuitem"`/`menuitemcheckbox` items, and each menu title reports its expanded state
- [ ] Toggle commands (ruler, outline pane, comments pane, comments-only view) render as checked items reflecting live state, replacing SPEC-037's standalone View menu without changing what those toggles do
- [ ] At a 1024px viewport the menu bar shows all six menus without scrolling, and the toolbar strip remains usable per SPEC-001

## Out of scope
The app bar and everything in it — the document title, the save-state indicator
and the Download menu SPEC-038 hardened stay exactly where they are; this spec
consolidates the toolbar strip only. User-customisable menus or toolbars. A
Word-style ribbon with tabs. The editor context menu (SPEC-016 owns it) and the
spell-check context menu (SPEC-035 owns it) — those are selection-targeted and
unaffected. Mobile and touch layout, still out of scope from SPEC-001. Adding,
removing or changing the behaviour of any command: this spec moves controls, it
does not alter what they do.

## Reference

### The contract this spec inherits, and the one it creates

**From `specs/001-app-shell-and-document-workspace.md`** — the shell is an app
bar, a toolbar strip below it, and a canvas. Its AC on the strip stands
unchanged and is not relitigated here:

> The toolbar strip is a horizontal container that scrolls horizontally when its
> contents overflow, rather than wrapping or clipping

This spec does not make the strip wrap, and does not remove its overflow
scrolling. It removes the *reason* the strip overflows. The 1280px criterion
above is about what is left in the strip, not a new layout mode for it.

**From `specs/021-keyboard-shortcuts-and-reference.md`** — `src/lib/shortcuts.ts`
is the one normative source for shortcuts. Menu items read their accelerator
label from it and never restate a binding:

```ts
export function shortcutLabel(id: string, isMac?: boolean): string
export const SHORTCUTS: ShortcutDefinition[]   // { id, label, category, bindings }
```

Its existing `category` values — File, Editing, Insert, Formatting, View,
Navigation, Find & replace — are why the menu set is File/Edit/View/Insert/
Format/Tools: five menus already have a category backing them. A command whose
menu placement and shortcut category disagree follows the menu; the category
field is for the SPEC-021 reference sheet, not for menu construction.

**From `specs/037-document-outline-navigator.md`** — the outline pane "can be
toggled from the View menu and its open/closed state persists per user". That
View menu becomes the View menu of this bar. The AC is satisfied by the
consolidated menu; `ViewMenu.vue` stops being a standalone toolbar dropdown.

**Created here, for every spec after this one:** a new command belongs in a
menu, and the toolbar strip is closed to additions unless the command is used
continuously enough to justify permanent screen space. A spec adding a command
names its menu. `specs/043-accessibility-checker.md` already does this
correctly ("Tools → Check accessibility") and needs no change once Tools exists.

### The mapping to build against

Grouped from the forty-three controls in the toolbar strip today. A spec merged
between this draft and its implementation adds its command to the matching
menu; the grouping principle is the deliverable, not this snapshot.

| Menu | Contents |
| --- | --- |
| File | Print, Page setup, Delete document |
| Edit | Undo, Redo, Cut, Copy, Paste, Paste without formatting, Select all, Find and replace, Autocorrect settings |
| View | Ruler, Outline pane, Comments pane, Comments-only view |
| Insert | Table, Image, Break, Header and footer, Footnote, Endnote, Table of contents, Special character, Link, Bookmark, Bookmark list, Comment |
| Format | Clear formatting, Format painter, Columns, Multilevel list, Text colour, Highlight colour, Paragraph spacing |
| Tools | Spelling and grammar, Document language, Custom dictionary, Word count |

Undo/Redo, text colour and highlight appear both in the retained toolbar and in
a menu. That is the one deliberate exception to the "exactly one menu" AC above,
which governs menus relative to each other — a control kept in the toolbar is
still listed in its menu so the menus stay a complete map.

### Not yet buildable

Nothing here is blocked. Every command listed is already implemented and already
in the toolbar; this spec relocates them. Tools is the only menu that starts
near-empty, and SPEC-035's language and dictionary controls plus SPEC-036's word
count fill it on arrival.
