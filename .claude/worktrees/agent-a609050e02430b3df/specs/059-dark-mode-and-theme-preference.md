---
id: SPEC-059
title: Dark mode and theme preference
status: draft
area: shell
depends_on: [SPEC-058]
---
## Why
SPEC-001 put "Theming or dark mode" out of scope in its first week, and
nothing has picked it up since. A word processor is an application people
keep open all day; every product it competes with — Google Docs,
LibreOffice Writer, Word — has shipped a dark interface, and a browser app
that ignores `prefers-color-scheme` is the one bright rectangle on an
otherwise dark desktop.

Dark mode in a document editor has one hard question the token layer does
not answer: **the page is not a surface**. A `.docx` with black text on a
white page is black text on a white page, and inverting it would be
lying about the document — but a full-brightness A4 sheet inside dark
chrome is exactly what makes naive dark modes unusable for writing. Word
and Writer both answer this the same way: chrome darkens, the page stays
paper, and a separate opt-in dims the page for on-screen reading without
touching the model. This spec takes that answer and states it precisely,
because the two things it must never affect are the saved file and the
printed output.

SPEC-058 made this cheap: every app colour is already a token on `:root`.
Dark mode is a second set of values for the same names, plus the handful
of places that paint colours from JavaScript rather than CSS — the canvas
painter's selection handles, the ruler, chart axes — which read tokens
through one helper instead of hard-coding.

## Acceptance criteria
- [ ] A theme preference with three values — `system`, `light`, `dark` — is settable from the View menu (`View → Appearance`) and persists across reloads under `docu-edit:theme`
- [ ] `system` follows `prefers-color-scheme` live: changing the OS appearance while the app is open re-themes it without a reload
- [ ] The chosen theme is applied as `data-theme="light" | "dark"` on `<html>`, set before first paint per "No flash of the wrong theme", and `system` sets no attribute — the media query alone decides
- [ ] `src/styles/tokens.css` gains a dark palette redefining **only** the token values, under both `@media (prefers-color-scheme: dark)` guarded as `:root:not([data-theme="light"])` and `:root[data-theme="dark"]`, so the explicit choice wins in both directions; no token is defined only inside a media or attribute block
- [ ] Every dark token pair meets the same WCAG AA contrast thresholds SPEC-058 asserts for light, in the same test, over both palettes
- [ ] The document page keeps its own colours in every theme — page background, text, table shading, image and shape colours come from the model and are never inverted, per "The page is not a surface"
- [ ] A separate, default-off `View → Dim page in dark theme` option renders the page at reduced luminance for on-screen reading only; it changes no model value, is not persisted into the document, and is ignored by print and PDF export
- [ ] The page's drop shadow, margin guides, page gaps, selection highlight, caret, and the workspace background around the pages are chrome and do follow the theme
- [ ] Colours painted from TypeScript — canvas selection outlines and handles, the anchor marker, table handles, ruler ticks and markers, chart gridlines and axis labels — read the resolved token values through `resolveThemeColor()` in "Theming what canvas paints", and repaint when the theme changes
- [ ] Print, print preview and PDF export are always light: no theme attribute reaches the print stylesheet and no dimming is applied, asserted by an e2e check that the exported PDF of a document is byte-identical in both themes
- [ ] Colour-carrying UI that shows *document* colour — the text-colour and highlight pickers, the shape fill/outline swatches, the chart palette — shows the true colour in both themes, over a neutral chip background so a white swatch stays visible on a dark surface
- [ ] Images, image thumbnails, template previews and the PDF preview canvas are never filtered or inverted
- [ ] Toasts, banners, dialog scrims and the drag-drop overlay are themed, and the scrim stays legible over both a dark chrome and a white page
- [ ] Under `prefers-reduced-motion: reduce` the theme switch is instantaneous; otherwise it may cross-fade over `--motion-base`
- [ ] An e2e test drives every top-level surface — library, editor, each of the seven panes, the menu bar, a dialog and the status bar — in dark theme and asserts no element resolves to a colour outside the dark palette

## Out of scope
User-authored or accent-colour themes, and any per-document appearance
setting — the theme is a property of the app on this device, not of the
file. A high-contrast or forced-colours mode (`forced-colors` media
queries) — worth its own spec. Changing the *document's* default text or
page colour, which is SPEC-077's theme-colour territory. Syncing the
preference across devices; there is no account (SPEC-001). Dark rendering
of the exported PDF.

## Reference

### The token layer this extends
`specs/058-design-system-and-visual-refresh.md` defines
`src/styles/tokens.css` as the only file in the repo holding a colour
literal, imported once from `App.vue`'s unscoped style block, with tokens
grouped as surfaces (`--surface-*`), content (`--text-*`), lines
(`--border-*`), accent/status, interaction overlays (`--state-*`,
`--focus-ring`), spacing, radius, shadow, type, chrome heights and motion.
This spec adds values, never names: if dark mode needs a colour that has no
token, the token is added to the light palette too.

The structure is the three-state one every themed page needs:

```css
:root { /* complete light palette — every token defined here */ }

@media (prefers-color-scheme: dark) {
  :root:not([data-theme="light"]) { /* dark values */ }
}

:root[data-theme="dark"] { /* the same dark values */ }
```

Defining a token *only* inside a media block is the bug this shape
prevents: a token that exists in one branch and not the other renders as
`unset` in the other.

### Dark palette — starting values
Surfaces are near-neutral dark greys, not pure black (pure black plus a
white page is the harshest possible pairing), and elevation goes *lighter*
as it comes forward — the inverse of the light palette's direction, which
is why the tokens are named by role and not by lightness.

```css
--surface-canvas:    #1f2023;   /* around the pages */
--surface-base:      #292a2d;   /* dialogs, panes, menus */
--surface-raised:    #35363a;   /* toolbar, status bar */
--surface-sunken:    #202124;
--surface-brand:     #202124;   /* the app bar — flat with the canvas */
--surface-selected:  #2d4159;
--surface-overlay:   rgb(0 0 0 / 0.6);
--text-primary:      #e8eaed;
--text-secondary:    #9aa0a6;
--text-disabled:     #5f6368;
--text-on-brand:     #e8eaed;
--text-accent:       #8ab4f8;
--text-danger:       #f28b82;
--border-subtle:     #35363a;
--border-default:    #5f6368;
--border-strong:     #9aa0a6;
--accent:            #8ab4f8;
--accent-hover:      #aecbfa;
--accent-subtle:     #1f3454;
--danger:            #f28b82;
--success:           #81c995;
--warning:           #fdd663;
--state-hover:       rgb(232 234 237 / 0.08);
--state-active:      rgb(232 234 237 / 0.14);
--focus-ring:        #8ab4f8;
--shadow-1: 0 1px 2px rgb(0 0 0 / 0.6);
--shadow-2: 0 2px 6px 1px rgb(0 0 0 / 0.5);
--shadow-3: 0 4px 12px 3px rgb(0 0 0 / 0.5);
```

Accent-on-dark inverts direction too: `--accent-hover` is *lighter* than
`--accent`, because a hover must add light on a dark surface. Any
primitive that hard-codes "hover = darken" is a SPEC-058 bug this spec
surfaces.

### The page is not a surface
The document is data. Its colours come from the model — `Run.color`,
`Paragraph`/`TableCell.shadingFill`, `Shape.fill`, `Image` pixels, and the
white of the paper itself — and a theme must not touch them, because what
you see must be what saves and what prints.

The boundary is exactly SPEC-058's exemption list. Inside
`.document-content` / `.pages`: model colours, untouched. Outside it, and
for the chrome *drawn around* pages — the workspace background, the page
drop shadow, the gap between pages, the dashed margin guides, the
non-printing-character marks, the ruler, the selection highlight and caret,
overlay handles — theme tokens apply.

**Dim page** is the one sanctioned exception, and it is a *view filter*:
the `.pages` container gets a reduced-luminance treatment (a CSS `filter`
or an overlay — the implementation's choice) applied at the container
level, never by rewriting any model value or any element's own colour. It
is off by default, available only while the resolved theme is dark, marked
"screen only" in the menu, persisted under `docu-edit:dim-page`, and
excluded from the print stylesheet by the existing `print-exclude`
mechanism the view already uses for on-screen-only chrome.

### No flash of the wrong theme
The stored preference must be applied before the first paint, so a
`<script>` in `index.html`'s `<head>` — ahead of the app bundle — reads
`localStorage["docu-edit:theme"]` and sets `data-theme` on
`documentElement` when the value is `light` or `dark`. It is a few lines,
it must not throw when storage is unavailable (SPEC-003's private-mode
case), and it is the only application logic allowed in `index.html`. The
Vue app reads the same key afterwards and owns every later change.

### Theming what canvas paints
`src/lib/**` is framework-free and must stay so — it cannot read CSS
variables, and `src/lib/core/**` is product-neutral besides (it is shared
with a future Slides host). So the *view* resolves colours and passes them
in, exactly as it already passes geometry:

```ts
// src/composables/useTheme.ts
export type ThemePreference = "system" | "light" | "dark";
export function useTheme(): {
  preference: Ref<ThemePreference>;          // persisted
  resolved: Readonly<Ref<"light" | "dark">>; // preference, with system resolved
  dimPage: Ref<boolean>;                     // persisted; only meaningful when dark
  setPreference(value: ThemePreference): void;
};

/** The computed value of a token on the current theme, e.g.
 *  resolveThemeColor("--accent") -> "#8ab4f8". Reads
 *  getComputedStyle(document.documentElement); returns "" when the token
 *  is unknown, so a caller's fallback is explicit. */
export function resolveThemeColor(token: string): string;
```

Painters that today embed a colour constant — the shape selection outline
and handles, the SPEC-051 anchor marker, `TableHandlesOverlay`, the ruler's
ticks and markers, chart gridlines and axis text (SPEC-052) — take the
colour as a parameter from their caller, which resolves it once per paint.
`resolved` changing invalidates the paginated render cache the same way a
page-setup change does, so a theme switch repaints the canvas; SPEC-047's
scheduler contract is unchanged.

Chart *series* colours are document data (SPEC-052's palette) and are not
themed. Chart *chrome* — gridlines, axis lines, tick labels, legend text —
is chrome and is.

### Menu placement
`View → Appearance` opens a submenu of three `menuitemradio` items
(System / Light / Dark) showing the current selection, followed by a
`menuitemcheckbox` for "Dim page in dark theme", disabled while the
resolved theme is light. Both register in the `editorCommands` Map in
`src/views/DocumentView.vue` — the one registry the menu bar and the
keyboard path both dispatch through (SPEC-049) — as `view.theme-system`,
`view.theme-light`, `view.theme-dark` and `view.dim-page`. No keyboard
shortcut is claimed.

The preference is app-level, not document-level, so it also applies to the
library (`HomeView.vue`) and the open-from-URL view, which have no menu
bar: they read the same composable and are themed by the same tokens.

### Persistence
Two keys, following the existing `docu-edit:*` convention used by
`docu-edit:outline-pane`, `docu-edit:autocorrect`,
`docu-edit:comments-pane` and friends:

- `docu-edit:theme` — `"system" | "light" | "dark"`; absent or
  unrecognised means `system`.
- `docu-edit:dim-page` — `"1"` when on; anything else off.

Both are written through the app's existing storage-availability guard, so
a blocked `localStorage` degrades to the session default rather than
throwing, per SPEC-003.
