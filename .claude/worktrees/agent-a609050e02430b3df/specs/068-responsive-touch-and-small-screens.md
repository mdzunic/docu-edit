---
id: SPEC-068
title: Responsive layout, touch and small screens
status: draft
area: shell
depends_on: [SPEC-058, SPEC-061, SPEC-063, SPEC-065]
---
## Why
"Mobile and touch layout" has been out of scope since SPEC-001 and was
restated as out of scope by SPEC-049. Sixty-seven specs later the product
is a browser application that cannot be used in a browser on the device
most people carry. It is not close, either: the menu bar, the toolbar and
the status bar together claim a fixed vertical budget; the page renders at
a fixed 96dpi width with no zoom until SPEC-063; every gesture is a mouse
gesture; and controls are sized for a pointer, not a fingertip.

The honest framing is that a phone cannot be a full word processor, and
pretending otherwise produces the worst version of both. What it can be is
a **capable reader and light editor**: open a document, read it at a
comfortable width, fix a typo, leave a comment, and download it. That is
the target this spec commits to, and it names what is deliberately not
offered on a small screen rather than shipping controls that do not fit.

Tablets are the case that pays best and is usually skipped: a 10-inch
tablet has the room for the real editor, and its problems are entirely
about touch — 32px targets, no hover, no right-click, text selection
handles, and a virtual keyboard that resizes the viewport under the caret.
Getting those right also improves the desktop build for anyone using a
touchscreen laptop.

## Acceptance criteria
- [ ] Three layout tiers are defined by available width, not by device sniffing — compact (<640px), medium (640–1023px) and full (≥1024px) — per "The tiers", and the app moves between them live as the window resizes
- [ ] At the full tier the layout is exactly what SPEC-058/061/065 specify; this spec changes nothing there
- [ ] At the medium tier the side panel overlays instead of displacing (SPEC-065), the toolbar collapses per SPEC-061's overflow order, and the menu bar stays a menu bar
- [ ] At the compact tier the menu bar collapses into a single "Menu" button opening the same six menus as a full-height sheet, and the toolbar reduces to the essential band in "The compact toolbar" plus an overflow
- [ ] Every command remains reachable at every tier — through the menu sheet, the overflow, or the command palette (SPEC-062) — and no command is silently dropped on a small screen
- [ ] A feature that genuinely cannot be operated by touch on a small screen is **disabled with a stated reason**, never missing: the affected commands are listed in "What compact does not offer", and each shows "not available on this screen size" in its menu item
- [ ] Interactive targets are at least 44×44 CSS px when the primary pointer is coarse (`@media (pointer: coarse)`), against 32px for a fine pointer, with the spacing scale adjusted to match rather than the icons enlarged
- [ ] Every hover-only affordance has a touch equivalent: tooltips do not carry unique information, the link chip appears on tap, table and drawing handles are visible while selected rather than on hover, and row overflow menus are always shown at the coarse-pointer tiers
- [ ] Right-click-only affordances have a touch equivalent: a long press opens the same context menu the editor, spell-check and table context menus provide, and the long-press threshold does not interfere with text selection
- [ ] Text selection by touch uses the browser's own selection with draggable handles; a tap places the caret, a double tap selects a word, a triple tap selects a paragraph, and the contextual text toolbar (SPEC-061) appears on selection end
- [ ] The virtual keyboard never covers the caret: the app uses the visual viewport to keep the caret in view when the keyboard opens, and restores the scroll position when it closes
- [ ] Pinch-to-zoom zooms the document through SPEC-063's view transform — not the browser's page zoom — and a two-finger pan scrolls; the app's own chrome does not scale
- [ ] At the compact tier the document defaults to Web layout at fit-width (SPEC-063), so text is readable without horizontal scrolling; Print layout remains selectable
- [ ] The library (SPEC-067) is usable at the compact tier: cards reflow to one column, the header actions collapse into a single primary action plus an overflow, and search stays visible
- [ ] Drag-based gestures that have no touch equivalent degrade to a dialog rather than being unreachable: dragging a ruler marker, a table boundary or a drawing handle each has a numeric-entry path
- [ ] Orientation changes and on-screen keyboard show/hide do not lose the caret, the selection, the scroll position or any open dialog's contents
- [ ] The app never sets `user-scalable=no` and never disables browser zoom, which would fail accessibility outright
- [ ] An e2e project runs the smoke path plus open/edit/comment/download at a 390×844 viewport with touch emulation, and the full suite continues to run at the desktop viewport unchanged

## Out of scope
A native app, an installable PWA, offline service workers, or any
app-store packaging. A separate mobile codebase or a second router. Stylus
and pressure input. Multi-touch drawing gestures beyond pinch and two-finger
pan. Landscape-specific layouts beyond what the width tiers already give.
Making every desktop feature work on a phone — the point of this spec is to
choose, and to say what was chosen. Performance work specific to mobile
hardware (SPEC-047 owns performance).

## Reference

### The tiers
Tiers key off the **available width of the app**, measured with a
`ResizeObserver` on the shell root, so a narrow window on a desktop
behaves identically to a tablet of the same width — the rule SPEC-061
already established for toolbar overflow.

| Tier | Width | Chrome |
| --- | --- | --- |
| Full | ≥1024px | Menu bar + toolbar + docked side panel + status bar |
| Medium | 640–1023px | Menu bar + collapsing toolbar + overlaid side panel + status bar |
| Compact | <640px | Menu button + essential toolbar + overlaid panel + reduced status bar |

Pointer coarseness is a **separate** axis, queried with
`@media (pointer: coarse)`, and controls target sizing only. A touchscreen
laptop at 1440px gets the full layout with large targets; a small window on
a mouse-driven desktop gets the compact layout with small ones. Conflating
the two is the standard mistake and this spec keeps them apart.

### The compact toolbar
One band, always visible: undo, redo, bold, italic, the style picker, and
the overflow button. Everything else lives in the overflow menu and the
menu sheet. The rationale matches SPEC-049's: what stays is what is used
continuously, and at 390px "continuously" is a much shorter list.

The status bar keeps the page indicator and the zoom control, and drops the
word count, the language control and the track-changes toggle into the menu
sheet.

### What compact does not offer
Disabled, with a reason, below 640px:

- Drawing shapes, text boxes and freeform (SPEC-051, SPEC-071) — creation
  gestures need a canvas that does not exist at this width. Existing
  drawings still render, and can be selected, moved and deleted.
- The chart data sheet (SPEC-052) — a spreadsheet grid at 390px is not
  usable. Charts render and can be moved.
- Mail merge (SPEC-046) and the document generation API surfaces.
- Print preview's multi-page layout (SPEC-034) — printing and PDF export
  themselves still work.
- Align and distribute (SPEC-072), and the drawing position dialog.

Each remains listed in its menu with the standard disabled reason, per
SPEC-049's rule that menus are a stable map of the product. Resizing the
window past 640px re-enables all of them with no reload.

### Touch equivalents for pointer gestures
| Desktop gesture | Touch equivalent |
| --- | --- |
| Hover tooltip | None; the label is on the control or in its menu item |
| Hover-revealed handles | Handles show while the object is selected |
| Right click | Long press (≈500ms), same menu |
| Click-drag ruler marker | Long press opens the numeric dialog; drag still works |
| Click-drag table boundary | Same, via Table properties |
| Click-drag drawing handle | Handles enlarge at coarse pointer; the position dialog is the fallback |
| Ctrl+wheel zoom | Pinch |
| Shift-click multi-select | Tap, then tap-with-Add from the contextual toolbar |

### The virtual keyboard
The app listens to `visualViewport`'s `resize` and `scroll` events and
keeps the caret rect inside the visible region, scrolling the document
rather than resizing it. The toolbar and status bar stay pinned to the
visible viewport, not the layout viewport, so they are not pushed behind
the keyboard. No layout is computed from `window.innerHeight` at the
compact tier — that is the value the keyboard invalidates.

### Testing
A second Playwright project runs a compact viewport (390×844) with
`hasTouch: true` and `isMobile: true`, covering: open a document from the
library, read it at fit-width, place the caret, type, apply bold from the
essential band, open the menu sheet and reach a Tools command, add a
comment, and download. Playwright stays pinned at 1.49.1 per CLAUDE.md —
this adds a project to the existing config, not a dependency.

`e2e/smoke.spec.ts` and `src/harness.spec.ts` are untouched. Existing
suites keep running at the desktop viewport, so no behavioural assertion
changes.

### Specs this changes
- `specs/001-app-shell-and-document-workspace.md` — "Mobile/touch layout"
  leaves its out-of-scope list; the shell's regions are unchanged at the
  full tier.
- `specs/049-menu-bar-and-toolbar-consolidation.md` — same: its
  out-of-scope note is resolved. The menu set, the placement of every
  command and the toolbar's contents at the full tier are untouched.
- `specs/061-toolbar-redesign-and-contextual-controls.md` — the compact
  band extends its overflow order; the order itself is unchanged.
- `specs/065-unified-side-panel.md` — its below-900px overlay rule is
  subsumed by the medium tier's boundary at 1024px; this spec's tier table
  is the normative one.
