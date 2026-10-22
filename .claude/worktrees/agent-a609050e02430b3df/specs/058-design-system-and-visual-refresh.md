---
id: SPEC-058
title: Design system — tokens, icons and the visual refresh
status: draft
area: shell
depends_on: [SPEC-001, SPEC-049]
---
## Why
Fifty-seven specs each styled their own component, and each picked its own
colours. `src/components/*.vue` and `src/views/*.vue` between them contain
**54 distinct hard-coded hex literals** and **zero CSS custom properties**.
They are not one palette drifting — they are two palettes fighting: the app
bar and its buttons use flat-UI blues and greys (`#2c3e50`, `#bdbdbd`,
`#e0e0e0`) inherited from SPEC-001, while every dialog, pane and menu built
since uses Material greys and blues (`#202124`, `#3c4043`, `#5f6368`,
`#dadce0`, `#f1f3f4`, `#1a73e8`). A user moving from the app bar to a
dialog crosses a visual seam, and a developer adding a control has no way
to know which of the two to copy — so they copy whichever file they had
open, and the count grows again.

The same is true of iconography, spacing and type. Only two components in
the repo contain an `<svg>` element (`ShapeGallery.vue`,
`ChartGallery.vue`); every other control is a text glyph — `B`, `I`, `U`,
`S`, `X₂`, `¶`, `ab` — set in the UI font at whatever size that file chose.
A toolbar of lettered buttons is the single largest reason the product
looks a decade older than it is. Padding is written per component in
whatever unit that component's author preferred (`0.25rem 0.75rem`,
`4px 8px`, `0.3em`), so nothing lines up across two components that were
never edited together.

This spec does not add a feature. It creates the vocabulary — tokens, an
icon set, a control primitive layer — that SPEC-059 through SPEC-068 are
written against, and migrates the existing surface onto it. Everything
after it stops choosing colours. The constraint that shapes the whole
design is CLAUDE.md's: `package.json` is protected, so there is no icon
package, no CSS framework and no design-token build tool. Everything here
is hand-authored files in `src/`.

## Acceptance criteria
- [ ] `src/styles/tokens.css` defines the complete token set in "Token vocabulary" below as CSS custom properties on `:root`, is imported once from `src/App.vue`'s global (unscoped) style block, and is the only file in the repo that contains a colour literal
- [ ] No `.vue` file under `src/` contains a hex colour, `rgb()`, `hsl()` or named CSS colour in a `<style>` block; a lint-style unit test (`src/styles/tokens.spec.ts`) reads every `.vue` file and fails on one, with the document-content and canvas exemptions named in "What is not a token"
- [ ] Every spacing, radius, border-width, shadow, font-size, font-weight and transition duration in a `.vue` `<style>` block resolves to a token; the same test enforces it for the properties listed in "Token vocabulary"
- [ ] The token values are chosen once, as one coherent light palette — the two inherited palettes are reconciled, not preserved side by side — and the app bar, menu bar, toolbar, dialogs, panes and status bar all read from it
- [ ] Every token colour pair used for text on a background meets WCAG AA contrast (4.5:1 body, 3:1 for text at or above 18.66px/bold 14px and for UI-component boundaries), asserted in `src/styles/tokens.spec.ts` by computing contrast from the token file itself
- [ ] `src/components/icons/` holds a hand-authored SVG icon for every toolbar and menu command that has one, exported through one `AppIcon.vue` component taking a `name` prop, per "The icon set"
- [ ] `AppIcon.vue` renders a 20×20 inline `<svg>` with `fill="none"`, `stroke="currentColor"`, `stroke-width="1.5"`, `aria-hidden="true"` and `focusable="false"`, so an icon inherits colour from its button and is never announced separately from the button's own accessible name
- [ ] An unknown icon name renders nothing and fails a unit test, so a typo cannot ship as an invisible button
- [ ] `src/components/ui/` holds the shared control primitives in "Control primitives" — `UiButton`, `UiIconButton`, `UiToggleButton`, `UiSelect`, `UiMenuSurface`, `UiField` — each with typed props, each token-styled, and the existing controls are migrated onto them
- [ ] `BaseDialog.vue` is the only dialog chrome in the repo: every dialog component renders inside it, with one header/body/footer structure, one focus-trap implementation and one Escape binding — per "Dialog consolidation"
- [ ] Focus is visible on every interactive control as a 2px `--focus-ring` outline with a 2px offset, never removed by `outline: none` without a replacement, asserted by the same style test
- [ ] Every control has a minimum 32×32px hit target at the default density, and the toolbar, menu bar and status bar each keep a single stated height token rather than sizing to content
- [ ] Hover, active, selected, disabled and focus states are defined once per primitive and applied by every consumer — no component re-implements a hover colour
- [ ] Motion is limited to the two transition tokens, and every transition is suppressed under `prefers-reduced-motion: reduce`
- [ ] The refreshed shell renders correctly at 1024px and 1280px viewport widths with no horizontal scrolling of app chrome, and the existing e2e suites pass unchanged except for assertions that name a colour or a glyph, per "Tests this changes"
- [ ] No dependency is added to `package.json`

## Out of scope
Dark mode (SPEC-059) — this spec defines the token *names* dark mode will
redefine, and one light palette. Any change to what a control does, where
it lives, or which menu holds it — SPEC-049 owns placement, SPEC-061 owns
the toolbar's contents. The document page itself: page background, text
rendering and the canvas painter keep their current appearance, because
they render the *document*, not the app (see "What is not a token").
Density modes / compact UI. User-selectable accent colours. Animation
beyond the two transition tokens. A component gallery or Storybook-style
page. Print and PDF output, which never sees app chrome.

## Reference

### The shell this restyles
`specs/001-app-shell-and-document-workspace.md` defines the shell and
`specs/049-menu-bar-and-toolbar-consolidation.md` the menu bar. As shipped,
`src/components/AppShell.vue` renders, in order: `header.app-bar`
(`data-testid="app-bar"`) holding `.product-name`, the `document-title`
h1/input pair, an optional `.format-badge` and an `app-bar-actions` slot;
`StorageWarningBanner`; a `.menubar-region` slot rendered only when a view
supplies one; `div.toolbar[role="toolbar"]`; `main.canvas`; a `.status-bar`
slot; then `ToastHost`, `ShortcutsDialog` and the drag-drop overlay. Those
regions and their `data-testid`s do not change here — only their styling.

### Token vocabulary — normative
`src/styles/tokens.css`, imported from `App.vue`'s unscoped block. Names
are the contract; the values below are the starting light palette and may
be tuned during implementation as long as the contrast AC holds.

```css
:root {
  /* Surfaces — ascending elevation, not ascending lightness. */
  --surface-canvas:        #f1f3f4;   /* behind the document pages */
  --surface-base:          #ffffff;   /* dialogs, panes, menus */
  --surface-raised:        #f8f9fa;   /* toolbar, status bar */
  --surface-sunken:        #e8eaed;   /* wells, inactive tabs */
  --surface-brand:         #1f2b3a;   /* the app bar */
  --surface-selected:      #e8f0fe;   /* selected row / active toggle */
  --surface-overlay:       rgb(32 33 36 / 0.45);   /* dialog scrim */

  /* Content — on-surface foregrounds. */
  --text-primary:          #202124;
  --text-secondary:        #5f6368;
  --text-disabled:         #9aa0a6;
  --text-on-brand:         #ffffff;
  --text-accent:           #1a73e8;
  --text-danger:           #b3261e;

  /* Lines. */
  --border-subtle:         #e8eaed;   /* separators inside a surface */
  --border-default:        #dadce0;   /* control outlines */
  --border-strong:         #9aa0a6;   /* emphasised/hovered outlines */

  /* Accent and status. */
  --accent:                #1a73e8;
  --accent-hover:          #1765cc;
  --accent-subtle:         #d2e3fc;
  --danger:                #b3261e;
  --success:               #188038;
  --warning:               #b06000;

  /* Interaction overlays, composited over any surface. */
  --state-hover:           rgb(32 33 36 / 0.06);
  --state-active:          rgb(32 33 36 / 0.12);
  --focus-ring:            #1a73e8;

  /* Spacing — a 4px scale; nothing between the steps. */
  --space-1: 2px;  --space-2: 4px;  --space-3: 8px;
  --space-4: 12px; --space-5: 16px; --space-6: 24px; --space-7: 32px;

  /* Radius, borders, elevation. */
  --radius-sm: 4px; --radius-md: 8px; --radius-lg: 12px; --radius-pill: 999px;
  --border-width: 1px;
  --shadow-1: 0 1px 2px rgb(60 64 67 / 0.2);
  --shadow-2: 0 2px 6px 1px rgb(60 64 67 / 0.15);
  --shadow-3: 0 4px 12px 3px rgb(60 64 67 / 0.15);

  /* Type — the UI font stack; the document has its own (SPEC-048). */
  --font-ui: system-ui, -apple-system, "Segoe UI", Roboto, sans-serif;
  --font-mono: ui-monospace, "SF Mono", Menlo, Consolas, monospace;
  --text-xs: 11px; --text-sm: 12px; --text-md: 13px;
  --text-lg: 15px; --text-xl: 18px;
  --weight-regular: 400; --weight-medium: 500; --weight-bold: 600;
  --line-tight: 1.2; --line-normal: 1.45;

  /* Chrome heights — one stated height per region, not content-sized. */
  --height-app-bar: 48px;
  --height-menu-bar: 32px;
  --height-toolbar: 40px;
  --height-status-bar: 28px;
  --height-control: 32px;      /* the minimum hit target */
  --width-pane: 280px;

  /* Motion. */
  --motion-fast: 120ms;
  --motion-base: 200ms;
  --easing: cubic-bezier(0.2, 0, 0, 1);
}
```

The token test enforces tokens for these declarations in `.vue` scoped
styles: `color`, `background`, `background-color`, `border*`, `outline*`,
`box-shadow`, `fill`, `stroke`, `padding*`, `margin*`, `gap`, `border-radius`,
`font-family`, `font-size`, `font-weight`, `line-height`, `transition*`.

### What is not a token
Three exemptions, each because the value is *document* data or geometry,
not app chrome. The test names them explicitly so the exemption cannot be
widened silently:

1. **Document content.** Anything under the `.document-content` /
   `.pages` subtree renders the user's document — its colours come from
   the model (`Run.color`, `TableCell.shadingFill`, `Shape.fill`) and its
   page white is the paper, not a surface. Styles scoped to those roots
   are exempt.
2. **The canvas painter.** `src/lib/**` is framework-free and paints
   shapes and charts with model colours; SPEC-051's
   `DEFAULT_NEW_SHAPE_FILL` (`#4472c4`) and friends are document defaults
   matching Word's, not UI colours. `src/lib/` is out of the test's scope
   entirely.
3. **Geometry in `px` that mirrors the model.** Ruler tick spacing,
   handle sizes and overlay rects derive from `PX_PER_CM`
   (`src/lib/core/units.ts`) and stay computed, not tokenised.

### The icon set
`src/components/icons/` holds one `.ts` module exporting the path data for
each icon, and `AppIcon.vue` renders it:

```vue
<!-- src/components/AppIcon.vue -->
<AppIcon name="bold" />        <!-- 20x20, currentColor, aria-hidden -->
```

Icons are drawn on a 20×20 grid, stroke-only (`fill="none"`,
`stroke="currentColor"`, `stroke-width="1.5"`, round caps and joins), so
one file serves every colour and state. The set covers, at minimum, every
command currently rendered as a text glyph: bold, italic, underline,
strikethrough, subscript, superscript; align left/center/right/justify;
bullet list, numbered list, multilevel list; indent increase/decrease;
line spacing; text colour, highlight; clear formatting, format painter;
undo, redo; cut, copy, paste; link, bookmark, comment, image, table,
chart, shape, text box, equation, special character, footnote, page break;
find, print, download, close; ruler, outline, styles, comments, track
changes, accessibility, fonts, history; zoom in, zoom out; chevrons
(up/down/left/right), check, dash, plus, minus, more-horizontal,
more-vertical, search, settings, warning, error, info.

Where a control's meaning is carried by its *letterform* — the `B`, `I`,
`U`, `S` of character formatting — the icon is the letterform drawn as a
glyph in the icon grid rather than typed into the button, so it sizes,
strokes and aligns with everything else.

Every icon-only control still carries a text `aria-label` and a tooltip
(SPEC-061 defines the tooltip's shortcut hint); the icon itself is always
`aria-hidden`.

### Control primitives
`src/components/ui/`, each an SFC with fully typed props and no feature
knowledge:

| Component | Purpose | Key props |
| --- | --- | --- |
| `UiButton` | Labelled button | `variant: "primary" \| "secondary" \| "ghost" \| "danger"`, `size: "sm" \| "md"`, `disabled` |
| `UiIconButton` | Icon-only button | `icon`, `label` (required — becomes `aria-label` and tooltip), `disabled` |
| `UiToggleButton` | Two-state control | `icon`/`label`, `pressed` (renders `aria-pressed`), `disabled` |
| `UiSelect` | Native `<select>` restyled | `modelValue`, `options`, `label`, `disabled` |
| `UiMenuSurface` | Popup container: elevation, radius, padding, outside-click and Escape dismissal, focus return | `open`, `anchor` |
| `UiField` | Label + control + optional hint/error for dialogs | `label`, `hint`, `error`, `for` |

`UiMenuSurface` is what stops each dropdown re-implementing dismissal;
`MenuBar.vue`, `DownloadMenu.vue`, `StylePicker.vue`, `HighlightPicker.vue`,
`TextColorPicker.vue`, `ShapeGallery.vue`, `ChartGallery.vue`,
`FontFamilySelect.vue` and `LanguageStatusControl.vue` all adopt it. Its
ARIA roles stay the consumer's choice — SPEC-049's `role="menu"` semantics
are unchanged, and a gallery stays a `listbox`/`grid` if that is what it is
today.

### Dialog consolidation
`BaseDialog.vue` already exists and is used by some dialogs. This spec
makes it universal: header (title + close `UiIconButton`), body (scrolls),
footer (right-aligned `UiButton`s, primary last), `role="dialog"`
`aria-modal="true"`, labelled by its title, focus trapped inside, Escape
closes, focus returns to the opener. Roughly thirty dialog components move
onto it. Where a dialog currently implements its own trap or Escape
handler, that code is deleted rather than left dormant.

### Tests this changes
The style/contrast/icon tests are new. Existing tests change only where
they assert what this spec deliberately alters:

- Assertions on a literal colour or on a glyph's text content (e.g.
  expecting a button's text to be `"B"`) become assertions on the
  control's accessible name or `data-testid`. Selecting by role and
  test-id is already the convention CLAUDE.md states, so most suites need
  no change at all.
- `src/harness.spec.ts` and `e2e/smoke.spec.ts` are untouched, per the
  testing rules.
- No `data-testid` is renamed or removed by this spec. A restyled control
  keeps its handle, so no e2e selector breaks.
- `DocumentView.size.spec.ts`'s budget only ratchets down: moving markup
  onto primitives removes lines from the view, and the budget is reset to
  the new count.

This is a refactor with a visual result, so the guard is that behaviour
tests keep passing unmodified — a suite that needed rewriting to stay
green would mean the refresh changed behaviour, which it must not.
