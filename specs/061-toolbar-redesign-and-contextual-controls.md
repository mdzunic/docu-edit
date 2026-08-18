---
id: SPEC-061
title: Toolbar redesign — icons, grouping, overflow and contextual controls
status: draft
area: shell
depends_on: [SPEC-049, SPEC-058]
---
## Why
SPEC-049 emptied the toolbar of the things people use twice a document and
left the continuously-used ones. It did not decide how those look, and the
result reads as one undifferentiated row of ~25 small bordered boxes:
undo/redo, the style picker, font family, a `− 11 +` size stepper, `B I U
S X₂ X²`, four alignment buttons, indents, lists, line spacing, paragraph
direction, pagination, text colour, highlight, and — for reasons no spec
defends — a `Track changes: Off` text button that also exists in the status
bar. Nothing is grouped, nothing is separated, and the character-formatting
controls are literal letters typed into buttons.

Two consequences. Visually, letterform buttons at UI-font size are the
single loudest "this is an unfinished web app" signal the product sends;
SPEC-058 builds the icon set that fixes it, and this spec is where the
toolbar actually adopts it. Structurally, SPEC-049's promise —
"no longer overflows at a 1280px viewport" — is a promise about *today's*
control count, and it is already tight. Every future spec that earns a
continuous-use control (SPEC-063's zoom, SPEC-072's align/distribute)
re-opens the question SPEC-049 closed. A toolbar with a defined overflow
behaviour answers it permanently: controls collapse in a stated priority
order instead of the strip scrolling sideways past the edge of the window.

The third gap is contextual. Selecting a shape, image, chart or field
already raises a purpose-built floating toolbar (`ShapeToolbar.vue`,
`ImageToolbar.vue`, `ChartToolbar.vue`, `FieldToolbar.vue`) — four
components, four layouts, four dismissal implementations, and nothing
equivalent for the most common selection of all, which is text.

## Acceptance criteria
- [ ] Every toolbar control renders as an icon from SPEC-058's `AppIcon` set, sized to `--height-control`, with a text `aria-label` and a tooltip; no control's meaning is carried by typed text except the font-family name and the font-size number
- [ ] Controls are grouped into the ordered bands in "Toolbar composition", separated by a 1px `--border-subtle` rule with `--space-3` on each side, and the bands' order is fixed rather than per-viewport
- [ ] A tooltip shows the command's label and, when one exists, its shortcut label read from `shortcutLabel(id)` in `src/lib/shortcuts.ts` — never a binding written out in the component
- [ ] Tooltips appear after a delay, are dismissed by Escape and by pointer-out, never trap focus, and are suppressed for a control whose accessible name is already visible
- [ ] When the strip cannot fit its bands, the lowest-priority bands collapse — whole band at a time, in the priority order in "Overflow" — into a trailing "More formatting" `UiMenuSurface` reached by a `more-horizontal` button, and the strip never scrolls horizontally
- [ ] Overflow is measured from the actual available width (a `ResizeObserver` on the strip), not from a viewport breakpoint, so opening a side pane collapses the toolbar exactly as narrowing the window does
- [ ] A command in the overflow menu is the same command with the same id, state and disabled logic — it is relocated, not duplicated in code
- [ ] Toggle controls render pressed state with `aria-pressed` and a `--surface-selected` background; a mixed selection (some bold, some not) renders an explicit indeterminate state, distinct from both on and off
- [ ] The duplicate `Track changes: Off` toolbar button is removed; the status-bar control (`status-bar-track-changes`) remains the single one, per "Removals"
- [ ] The font-size stepper becomes one control — an editable numeric field with a dropdown of common sizes and `−`/`+` steppers — rather than three adjacent buttons
- [ ] The four floating object toolbars are unified onto one `ContextualToolbar.vue` chrome — positioning, collision handling against the viewport and panes, dismissal, focus return and ARIA — with each object kind supplying only its own controls
- [ ] Selecting a non-empty range of text raises a contextual text toolbar with bold, italic, underline, text colour, highlight, link and comment, positioned above the selection, per "The text selection toolbar"
- [ ] The text selection toolbar never appears for a collapsed caret, during a drag, while a menu or dialog is open, or while the object toolbars are showing, and is dismissed by Escape without clearing the selection
- [ ] Every toolbar control is reachable by keyboard: the strip is a single tab stop with Left/Right arrow roving focus (the `role="toolbar"` convention), Home/End jump to the ends, and the overflow button is the last stop
- [ ] The toolbar keeps its stated `--height-toolbar` at every width, and at 1024px shows at least the character- and paragraph-formatting bands without overflow
- [ ] Existing e2e selectors keep working: no `data-testid` on a retained control is renamed, and a control that moves into the overflow menu is still found by the same test id once the menu is opened

## Out of scope
Which commands exist and which menu holds them — SPEC-049 owns that, and
this spec moves nothing between menus. User-customisable toolbars, or a
ribbon. Adding any new command. The menu bar's own visual design beyond
what SPEC-058's tokens give it. The status bar's contents (SPEC-063 adds
zoom there). Touch-sized targets and small-screen layout (SPEC-068).
Changing the editor context menu (SPEC-016) or the spell-check context menu
(SPEC-035) — both are selection-targeted and stay as they are.

## Reference

### What ships today
`src/views/DocumentView.vue` fills `AppShell`'s `toolbar` slot with, in
order: `UndoRedoControls`, `StylePicker`, `FontFamilySelect`,
`FontSizeControl`, `FormatToolbar`, `FieldToolbar`, `ParagraphToolbar`,
`TextColorPicker`, `HighlightPicker`, and an inline
`data-testid="toolbar-track-changes-toggle"` button whose title comes from
`shortcutLabel('edit.toggleTrackChanges')`.

`ParagraphToolbar.vue`'s controls carry the testids `list-bulleted`,
`list-bulleted-menu-button`, `bullet-style-menu`, `list-numbered`,
`list-numbered-menu-button`, `number-style-menu`, `indent-increase`,
`indent-decrease`, `line-spacing-button`, `line-spacing-menu`,
`line-spacing-custom`, `space-before-toggle`, `paragraph-direction`,
`paragraph-pagination-button`, `paragraph-pagination-menu`. Those names are
the e2e contract and survive this spec.

The floating toolbars are `ShapeToolbar.vue` (`shape-toolbar`, plus
`shape-fill-color`, `shape-outline-*`, `shape-rotation`, `shape-wrap`,
`shape-zorder`, `shape-group-button`, `shape-ungroup-button`,
`shape-alt-text`, `shape-duplicate-button`, `shape-delete-button`),
`ImageToolbar.vue`, `ChartToolbar.vue` and `FieldToolbar.vue`
(`field-toolbar`, `field-toolbar-edit|update|lock|unlink|toggle-codes`).

### Toolbar composition
Bands, in fixed order, left to right:

| # | Band | Controls |
| --- | --- | --- |
| 1 | History | undo, redo |
| 2 | Style | paragraph style picker |
| 3 | Type | font family, font size |
| 4 | Character | bold, italic, underline, strikethrough, subscript, superscript |
| 5 | Colour | text colour, highlight |
| 6 | Paragraph | align left/center/right/justify, line spacing, space before |
| 7 | Lists | bulleted, numbered, indent decrease, indent increase |
| 8 | Paragraph extras | direction, pagination options |

`FieldToolbar` leaves the strip entirely — it is contextual, shown only
while the caret is in a field, and belongs on `ContextualToolbar` with the
other three.

### Overflow
Bands collapse from the **highest number down**: 8, then 7, then 6, then 5.
Bands 1–4 never collapse — history, style, type and character formatting
are the controls that justify permanent screen space, which is SPEC-049's
own test for what stays in the strip. A collapsed band renders inside the
overflow menu under its band name as a section heading, keeping the same
grouping the strip shows.

Measurement is by `ResizeObserver` on the strip element, comparing the sum
of band widths (measured once after fonts settle, cached) against available
width. Hysteresis of one band width prevents flapping at the boundary.

### The text selection toolbar
Raised when the selection is a non-empty range **within body text**:

- Position: horizontally centred on the selection's bounding rect,
  vertically above it, flipped below when it would leave the viewport or
  overlap the toolbar strip; it never covers the selection itself.
- Contents: bold, italic, underline, text colour, highlight, link,
  comment. Deliberately short — it is a shortcut to the seven things
  people reach for mid-sentence, not a second toolbar.
- Suppressed while: the selection is collapsed; a drag is in progress
  (it appears on pointer-up); a dialog, menu or the find bar is open; an
  object is selected; shape text is being edited; the document is
  read-only; track-changes preview is showing.
- Escape hides it and leaves the selection alone. It is not a tab stop —
  keyboard users reach the same commands from the strip and the menus,
  and a floating tab stop mid-document is a worse experience than not
  having one.

It shares `ContextualToolbar.vue`'s chrome with the object toolbars, so
positioning and dismissal exist once.

### Removals
Exactly one control is removed: the toolbar's `Track changes: Off` button.
It duplicates `status-bar-track-changes`, which SPEC-039 also ships, and
the keyboard command (`edit.toggleTrackChanges`) and its menu item both
remain. `e2e` assertions that target `toolbar-track-changes-toggle` move to
the status-bar control; this is the sanctioned narrow case of a test
changing because the spec deliberately removed the thing it asserted, and
it is bounded to that one testid.

### Accessibility
The strip keeps `role="toolbar"` with `aria-label="Document toolbar"`
(SPEC-001). Roving `tabindex` makes it one tab stop, per the ARIA toolbar
pattern; Left/Right move, Home/End jump, and a control that opens a menu
delegates arrow keys to that menu while it is open. Icon-only buttons
carry `aria-label`; their `AppIcon` is `aria-hidden` (SPEC-058), so nothing
is announced twice. Indeterminate toggles report `aria-pressed="mixed"`.
