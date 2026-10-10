// SPEC-058 — hand-authored icon geometry. Every icon is drawn on a 20×20 grid
// and rendered stroke-only by AppIcon.vue (fill="none", stroke="currentColor",
// stroke-width 1.5, round caps/joins), so one path set serves every colour and
// state — the icon always inherits its button's `color`.
//
// Each entry is the list of SVG path `d` strings that make up the glyph.
// Letterform commands (B, I, U, S, x², x₂) are drawn as strokes on the same
// grid rather than typed as text, so they size and align with every other
// icon. Keep names in sync with AppIcon's `name` prop and the toolbar/menu
// call sites; tokens.spec.ts / AppIcon.spec.ts assert an unknown name renders
// nothing, so a typo cannot ship as an invisible button.
export const ICON_PATHS: Record<string, string[]> = {
  // ---- Character formatting (letterforms drawn on the grid) ----------------
  bold: [
    "M6 4h5a3 3 0 0 1 0 6H6z",
    "M6 10h6a3 3 0 0 1 0 6H6z",
  ],
  italic: ["M12 4h-3", "M11 16H8", "M11.5 4l-3 12"],
  underline: ["M6 4v5a4 4 0 0 0 8 0V4", "M5 17h10"],
  strikethrough: [
    "M7 6a3 3 0 0 1 3-2h1a3 3 0 0 1 3 2.5",
    "M13 12a3 3 0 0 1-3 4H9a3 3 0 0 1-3-2.5",
    "M4 10h12",
  ],
  subscript: ["M4 5l7 8", "M11 5l-7 8", "M18 16h-3l3-3v-.5a1.2 1.2 0 0 0-2.4 0"],
  superscript: ["M4 7l7 8", "M11 7l-7 8", "M18 8h-3l3-3v-.5a1.2 1.2 0 0 0-2.4 0"],

  // ---- Paragraph alignment / lists ----------------------------------------
  "align-left": ["M4 5h12", "M4 9h8", "M4 13h12", "M4 17h8"],
  "align-center": ["M4 5h12", "M6 9h8", "M4 13h12", "M6 17h8"],
  "align-right": ["M4 5h12", "M8 9h8", "M4 13h12", "M8 17h8"],
  "align-justify": ["M4 5h12", "M4 9h12", "M4 13h12", "M4 17h12"],
  "bullet-list": ["M8 6h9", "M8 10h9", "M8 14h9", "M4 6h.01", "M4 10h.01", "M4 14h.01"],
  "numbered-list": [
    "M9 6h8", "M9 10h8", "M9 14h8",
    "M4 5v3", "M4 12h1.5l-1.5 2h1.5",
  ],
  "multilevel-list": [
    "M4 6h2", "M8 6h9", "M8 10h9", "M11 14h6", "M7 14h1",
  ],
  "indent-increase": ["M4 5h12", "M9 9h7", "M9 13h7", "M4 17h12", "M4 8v6l3-3z"],
  "indent-decrease": ["M4 5h12", "M9 9h7", "M9 13h7", "M4 17h12", "M7 8v6l-3-3z"],
  "line-spacing": ["M9 6h8", "M9 10h8", "M9 14h8", "M4 5v10", "M4 5l-1.5 1.5", "M4 5l1.5 1.5", "M4 15l-1.5-1.5", "M4 15l1.5 1.5"],
  // Paragraph direction (RTL): a paragraph mark with a left-pointing flow
  // arrow beneath, drawn on the grid rather than typed as "¶⇄".
  "paragraph-direction": ["M11 5v10", "M8 5v10", "M8 5a3 3 0 0 0 0 6h3", "M13 5h-5", "M9 17H4", "M4 17l2-2", "M4 17l2 2"],
  // Paragraph pagination (keep-together / break control): a page outline with
  // a paragraph mark inside, drawn rather than typed as "¶⇟".
  "paragraph-pagination": ["M5 4h8l2 2v10H5z", "M13 4v2h2", "M11 8H8.5a1.5 1.5 0 0 0 0 3H11", "M11 8v6", "M9.5 8v6"],

  // ---- Colour / painter ----------------------------------------------------
  "text-color": ["M5 14l3.5-9h1L13 14", "M6.5 11h5", "M4 17h12"],
  highlight: ["M4 16h5", "M6 13l6-6 3 3-6 6H6z", "M11 5l2-2 3 3-2 2"],
  "clear-formatting": ["M6 5h9", "M11 5l-3 9", "M4 17h6", "M13 12l4 4", "M17 12l-4 4"],
  "format-painter": ["M4 5h9v3H4z", "M13 6.5h2v3H9v2", "M8 11h3v6H8z"],

  // ---- Sheet formatting (SPEC-123) -------------------------------------------
  // A paint bucket tipping onto a swatch bar; the bucket's lip is the spout.
  "fill-color": ["M5 9l5-5 5 5-5 5z", "M15 9l1.5 3a1.5 1.5 0 0 1-3 0z", "M4 17h12"],
  // A cell grid with its outer box drawn heavier by a second stroke inset.
  borders: ["M4 4h12v12H4z", "M4 10h12", "M10 4v12", "M6 6h8v8H6z"],
  "align-top": ["M4 4h12", "M10 7v9", "M7 10l3-3 3 3"],
  "align-middle": ["M4 10h12", "M10 4v3", "M10 13v3", "M8 5l2 2 2-2", "M8 15l2-2 2 2"],
  "align-bottom": ["M4 16h12", "M10 4v9", "M7 10l3 3 3-3"],
  // Text lines wrapping back round at the right edge.
  "wrap-text": ["M4 5h12", "M4 10h9a2.5 2.5 0 0 1 0 5h-3", "M12 13l-2 2 2 2", "M4 15h3"],
  // An "ab" baseline tilted up, with the angle's arc.
  "text-rotation": ["M5 15l7-7", "M9 8l3 0 0 3", "M5 9a6 6 0 0 1 6 6", "M4 16h12"],
  currency: ["M10 4v12", "M13 7a3 3 0 0 0-3-1.5c-1.7 0-3 1-3 2.3S8.3 9.6 10 10s3 1 3 2.2-1.3 2.3-3 2.3a3 3 0 0 1-3-1.5"],
  percent: ["M15 5L5 15", "M7 5.5a1.5 1.5 0 1 1-1.5 1.5A1.5 1.5 0 0 1 7 5.5z", "M13 11.5a1.5 1.5 0 1 1-1.5 1.5 1.5 1.5 0 0 1 1.5-1.5z"],
  // A thousands comma between digit marks.
  "comma-style": ["M6 6v8", "M14 6v8", "M10 12.5v1.5a1.5 1.5 0 0 1-1.5 1.5"],
  "decimals-increase": ["M4 14h.01", "M7 9v5", "M10 9v5", "M13 9v5", "M15 5h4", "M17 3v4"],
  "decimals-decrease": ["M4 14h.01", "M7 9v5", "M10 9v5", "M15 5h4"],
  // A stack of three style chips.
  "cell-styles": ["M4 4h12v4H4z", "M4 10h12v3H4z", "M4 15h7v2H4z"],
  // Two cells joined under one span arrow (SPEC-124's Merge & Centre).
  "merge-cells": ["M3 5h14v10H3z", "M6 10h8", "M8 8l-2 2 2 2", "M12 8l2 2-2 2"],

  // ---- History -------------------------------------------------------------
  undo: ["M7 7L4 10l3 3", "M4 10h8a4 4 0 0 1 0 8H9"],
  redo: ["M13 7l3 3-3 3", "M16 10H8a4 4 0 0 0 0 8h3"],

  // ---- Clipboard -----------------------------------------------------------
  // Symmetric finger loops (the old loops sat at 4,14 and 14,14 — asymmetric;
  // corrected geometry from the library-modernist design handoff).
  cut: ["M6.9 13.6L14.5 3.5", "M13.1 13.6L5.5 3.5", "M6.6 13.4a2.2 2.2 0 1 0 0 4.4 2.2 2.2 0 0 0 0-4.4z", "M13.4 13.4a2.2 2.2 0 1 0 0 4.4 2.2 2.2 0 0 0 0-4.4z"],
  copy: ["M7 7h8v9H7z", "M5 13H4V4h9v1"],
  paste: ["M6 5h8v11H6z", "M8 5V4h4v1", "M8 9h4", "M8 12h4"],

  // ---- Insert --------------------------------------------------------------
  link: ["M9 11a3 3 0 0 0 4 0l2-2a3 3 0 0 0-4-4l-1 1", "M11 9a3 3 0 0 0-4 0l-2 2a3 3 0 0 0 4 4l1-1"],
  bookmark: ["M6 4h8v13l-4-3-4 3z"],
  comment: ["M4 5h12v8H9l-3 3v-3H4z"],
  image: ["M4 5h12v10H4z", "M4 13l4-4 3 3 3-3 2 2", "M8 8a1 1 0 1 1-1-1 1 1 0 0 1 1 1z"],
  table: ["M4 5h12v10H4z", "M4 9h12", "M4 13h12", "M9 5v10"],
  chart: ["M5 4v12h11", "M8 13v-3", "M11 13V8", "M14 13v-5"],
  shape: ["M6 5h8v10H6z"],
  "text-box": ["M4 6h12v8H4z", "M7 9h6", "M10 9v4"],
  equation: ["M5 5h9l-4 5 4 5H5", "M8 10h4"],
  "special-character": ["M7 15V8a3 3 0 0 1 6 0v7", "M5 15h4", "M11 15h4"],
  footnote: ["M5 5h7", "M5 9h7", "M5 13h4", "M13 11h3l-3 4h3"],
  "page-break": ["M6 4h5l3 3v3H6z", "M6 16v-3h8v3", "M4 12h12"],

  // ---- Actions -------------------------------------------------------------
  // Lens centred at 9,9 so the handle leaves the rim tangentially (the old arc
  // and handle were centred inconsistently; corrected geometry from the
  // library-modernist design handoff — zoom-in/out below share it).
  find: ["M9 4.5a4.5 4.5 0 1 0 0 9 4.5 4.5 0 0 0 0-9z", "M12.5 12.5l4 4"],
  search: ["M9 4.5a4.5 4.5 0 1 0 0 9 4.5 4.5 0 0 0 0-9z", "M12.5 12.5l4 4"],
  // Paper tray abuts the body outline instead of crossing it (handoff fix).
  print: ["M6 8V4h8v4", "M4 8h12v5h-2v-2H6v2H4z", "M6 13h8v4H6z"],
  download: ["M10 4v8", "M6 9l4 4 4-4", "M5 16h10"],
  close: ["M6 6l8 8", "M14 6l-8 8"],

  // ---- View / panes --------------------------------------------------------
  ruler: ["M4 7h12v6H4z", "M7 7v2", "M10 7v3", "M13 7v2"],
  outline: ["M5 6h2", "M9 6h6", "M5 10h2", "M9 10h6", "M5 14h2", "M9 14h6"],
  styles: ["M5 5h6v4H5z", "M13 5h2v10h-2z", "M5 11h6v4H5z"],
  comments: ["M4 5h12v8H9l-3 3v-3H4z", "M7 9h6"],
  // Ferrule stroke inside the open barrel (handoff fix; pencil below shares
  // the barrel).
  "track-changes": ["M4.5 15.5l1-3 7.9-7.9 2 2-7.9 7.9z", "M11.4 6.6l2 2", "M4.5 17.5h7"],
  accessibility: ["M10 6a1 1 0 1 1-1-1 1 1 0 0 1 1 1z", "M5 8h10", "M10 8v4", "M7 16l3-4 3 4"],
  fonts: ["M5 15l4-10 4 10", "M6.5 12h5", "M14 15h2"],
  history: ["M10 6a5 5 0 1 1-5 5", "M5 7v3h3", "M10 8v3l2 1"],

  // ---- Zoom ----------------------------------------------------------------
  // Lens and ± bars share the 9,9 centre (the old lens arc was centred at 5,5
  // while the bars were centred at 7,7 — handoff fix).
  "zoom-in": ["M9 4.5a4.5 4.5 0 1 0 0 9 4.5 4.5 0 0 0 0-9z", "M12.5 12.5l4 4", "M6.5 9h5", "M9 6.5v5"],
  "zoom-out": ["M9 4.5a4.5 4.5 0 1 0 0 9 4.5 4.5 0 0 0 0-9z", "M12.5 12.5l4 4", "M6.5 9h5"],

  // ---- Appearance (the Preferences theme switch) ---------------------------
  sun: [
    "M10 6.5a3.5 3.5 0 1 1 0 7 3.5 3.5 0 0 1 0-7z",
    "M10 3v1.5", "M10 15.5V17", "M3 10h1.5", "M15.5 10H17",
    "M5.05 5.05l1.06 1.06", "M13.89 13.89l1.06 1.06",
    "M14.95 5.05l-1.06 1.06", "M6.11 13.89l-1.06 1.06",
  ],
  moon: ["M13 4a7 7 0 1 0 3 8 5.5 5.5 0 0 1-3-8z"],
  // A display — "follow the system".
  system: ["M4 5h12v8H4z", "M8 16h4", "M10 13v3"],

  // ---- Chevrons & primitives ----------------------------------------------
  "chevron-up": ["M5 12l5-5 5 5"],
  "chevron-down": ["M5 8l5 5 5-5"],
  "chevron-left": ["M12 5l-5 5 5 5"],
  "chevron-right": ["M8 5l5 5-5 5"],
  check: ["M5 10l3 3 7-7"],
  dash: ["M5 10h10"],
  plus: ["M10 5v10", "M5 10h10"],
  minus: ["M5 10h10"],
  "more-horizontal": ["M5 10h.01", "M10 10h.01", "M15 10h.01"],
  "more-vertical": ["M10 5v.01", "M10 10v.01", "M10 15v.01"],
  settings: ["M10 8a2 2 0 1 1-2 2 2 2 0 0 1 2-2z", "M10 3v2", "M10 15v2", "M3 10h2", "M15 10h2", "M5 5l1.5 1.5", "M13.5 13.5L15 15", "M15 5l-1.5 1.5", "M6.5 13.5L5 15"],
  // The rename affordance on the app bar's document title (AppShell.vue).
  pencil: ["M4.5 15.5l1-3 7.9-7.9 2 2-7.9 7.9z", "M11.4 6.6l2 2"],
  warning: ["M10 4l6 11H4z", "M10 9v3", "M10 14h.01"],
  error: ["M10 4a6 6 0 1 1-6 6 6 6 0 0 1 6-6z", "M7.5 7.5l5 5", "M12.5 7.5l-5 5"],
  info: ["M10 4a6 6 0 1 1-6 6 6 6 0 0 1 6-6z", "M10 9v4", "M10 6.5h.01"],
};

export type IconName = keyof typeof ICON_PATHS;

export function hasIcon(name: string): boolean {
  return Object.prototype.hasOwnProperty.call(ICON_PATHS, name);
}
