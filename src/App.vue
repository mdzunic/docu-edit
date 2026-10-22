<script setup lang="ts"></script>

<template>
  <RouterView />
</template>

<style>
/* The design system's tokens (SPEC-058). Imported once here, from the app's
   only unscoped style block, so :root custom properties are visible to every
   scoped component style. Must be the first rule — @import is only valid at
   the top of a stylesheet. */
@import "./styles/tokens.css";

:root {
  font-family: var(--font-ui);
  color: var(--text-primary);
}

/* Form controls do not inherit color — the UA gives them `buttontext`, which
   stays near-black whatever the tokens say. Every control that sets a token
   background but no color (dozens of dialog/pane/toolbar rules) was therefore
   black-on-dark in the dark theme (SPEC-059). Inheriting makes the tokenised
   text colour the default everywhere; a control that needs a different colour
   still declares its own. */
button,
input,
select,
textarea {
  color: inherit;
}

/* The document page is paper: it keeps its own light colours in every theme
   (SPEC-059 AC-12), so UA-rendered widget chrome inside a page must never
   flip dark with the app. */
.page {
  color-scheme: light;
}

/* One focus ring for the whole app (SPEC-058 AC-45/46/56): every interactive
   control shows a 2px --focus-ring outline with a 2px offset. Written with
   :where() so its specificity is zero — a component's own :focus-visible rule
   can still refine it, but a control that adds no focus styling of its own
   (or only a text-decoration) inherits the correct ring instead of the
   browser's 1px default. --space-1 is 2px. */
:where(a[href], button, input, select, textarea, summary, [tabindex]:not([tabindex="-1"])):focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: var(--space-1);
}

html,
body,
#app {
  height: 100%;
  margin: 0;
}

/* Theme cross-fade (SPEC-059 AC-46): when motion is allowed, the top-level
   chrome surfaces ease their colours over --motion-base as the theme flips.
   Under prefers-reduced-motion the whole rule is dropped, so the switch is
   instantaneous (AC-45) — belt-and-braces with tokens.css already collapsing
   --motion-base to 0ms there. Scoped to shell surfaces so it never animates a
   component's own hover/active state. */
@media (prefers-reduced-motion: no-preference) {
  .app-bar,
  .toolbar,
  .menu-bar,
  .status-bar,
  .canvas {
    transition:
      background-color var(--motion-base) var(--easing),
      color var(--motion-base) var(--easing),
      border-color var(--motion-base) var(--easing);
  }
}

/* ---- Dim page for on-screen reading in dark theme (SPEC-059 AC-13) --------
   A container-level luminance reduction on the page boxes — and on a
   workbook's cell area, the sheet paper (SPEC-117 "Dark mode") — applied
   only when the resolved theme is dark AND the opt-in is on. It touches no model value
   and is not persisted into the document (AC-15). It is screen-only: the
   print stylesheet and print preview reset it (AC-16/32/73), so print, print
   preview and PDF export never dim — the same on-screen-only exclusion the
   view already uses for its chrome. */
:root[data-theme="dark"][data-dim-page="1"] .page,
:root[data-theme="dark"][data-dim-page="1"] .sheet-paper {
  filter: brightness(0.82);
}

body.is-print-preview .page {
  filter: none !important;
}

/* ---- Print & print preview (SPEC-034) ------------------------------------
   Screen-only affordances carry the `print-exclude` class on their real
   rendered root elements (ruler, find bar, table/image handles, the TOC
   update bar, popovers…): they hide both while the in-app print preview is
   open and in printed output (AC-13). Pages outside the selected range are
   marked data-print-hidden by DocPage and hide the same way (AC-29/31).
   The selectors below intentionally target the classes components actually
   render — e.g. the ruler's root class is `ruler` via `print-exclude`. */
body.is-print-preview .print-exclude {
  display: none !important;
}

body.is-print-preview .page[data-print-hidden] {
  display: none !important;
}

/* The margin guide is a screen affordance too: the preview shows the page
   as it prints. */
body.is-print-preview .page-margin {
  outline: none !important;
}

/* Field shading (SPEC-053 AC-12) is screen-only: the preview shows the page
   as it prints, so the grey field background is removed here too. */
body.is-print-preview .doc-field-shaded {
  background: transparent !important;
}

@media print {
  .print-exclude,
  .page[data-print-hidden] {
    display: none !important;
  }

  /* SPEC-104: printing a deck shows its off-screen print sheet (teleported
     to <body>) and nothing else, back on the page in normal flow. */
  body.is-deck-print > :not(.deck-print-sheet) {
    display: none !important;
  }

  body.is-deck-print .deck-print-sheet {
    position: static !important;
    left: auto !important;
  }

  /* Page backgrounds are opt-in on paper (SPEC-074 AC-25), matching both
     source applications' "print background colours" setting. The page
     BORDER has no such switch — it always prints. */
  .page-background {
    display: none !important;
  }

  body.prints-background .page-background {
    display: block !important;
  }

  /* The aria-live announcement stubs (.visually-hidden) are 1px absolutely
     positioned boxes whose static position is just past the last page box.
     That one pixel of overflow rounded the sheet count up, printing an empty
     trailing page. Printed output has no assistive-technology listeners. */
  .visually-hidden {
    display: none !important;
  }

  /* App chrome never prints: only the page boxes reach the sheets. The menu
     bar (SPEC-049) postdates this list and used to leak: its 33px row was the
     only flow content before the first page box, so it printed as a sheet of
     its own — "File Edit View…" on an otherwise empty first page. */
  .app-bar,
  .menubar-region,
  .toolbar,
  .status-bar,
  .storage-warning,
  .toast-region,
  .notice-bar,
  .drop-overlay,
  .drop-error,
  .print-preview-panel {
    display: none !important;
  }

  /* Unclamp the shell so the full page list flows onto print sheets
     instead of staying inside the screen scroll container. */
  html,
  body,
  #app {
    height: auto !important;
    overflow: visible !important;
  }

  .shell {
    display: block !important;
    height: auto !important;
    overflow: visible !important;
  }

  .canvas {
    height: auto !important;
    overflow: visible !important;
    background: none !important;
  }

  .page-wrap {
    padding: 0 !important;
  }

  /* One sheet per page box: the injected @page rule (usePrintPreview) has
     zero margin and the document's own page size, and the box's padding
     carries the document margins — so the sheet equals the on-screen page. */
  .page {
    margin: 0 auto !important;
    box-shadow: none !important;
    break-after: page;
    page-break-after: always;
  }

  .page:last-child {
    break-after: auto;
    page-break-after: auto;
  }

  .page-margin {
    outline: none !important;
  }

  /* The dim-page treatment is on-screen only: a sheet always prints at full
     luminance regardless of the theme or the dim opt-in (SPEC-059 AC-16/32). */
  .page {
    filter: none !important;
  }
}

/* SPEC-063 Focus mode: Print layout with the menu bar, toolbar and the open
   side panel hidden purely by CSS — the panel's own state is untouched, so the
   exact pane open on entry returns on exit (AC-44, AC-47). The status bar (page
   indicator + zoom control) stays visible (AC-45, AC-56). Global, not scoped,
   because it spans AppShell's chrome and DocumentView's side panel. SPEC-065
   consolidated every pane dock into one `.side-panel-root`; `!important` beats
   that component's scoped-class specificity tie so the open pane is hidden. */
.is-focus-mode .menubar-region,
.is-focus-mode .toolbar {
  display: none;
}

.is-focus-mode .side-panel-root {
  display: none !important;
}

/* SPEC-063 Web layout: the same paginated model rendered as one continuous
   column — page frames, gaps, headers, footers and page numbers hidden
   (AC-38, AC-39). The model and section properties are untouched (AC-42).

   Every override below is SCREEN-ONLY and additionally suppressed while the
   print preview is open or the PDF export is reading the DOM. AC-52 says
   export and print always see 100% Print layout whatever the view mode, and
   `width`/`padding` cannot be handed back by `revert` — an `!important`
   author rule beats DocPage's inline `width: <pageWidth>cm`, and reverting
   drops the inline value too. Not applying the rules at all is what restores
   the real sheet. Without this the exporter mapped a content-width page box
   onto an A4 media box: every font size grew by the ratio (12pt printed as
   15.8pt) and the margins vanished. */
@media screen {
  body:not(.is-print-preview):not(.is-exporting-pdf) .is-web-layout .page {
    /* !important throughout: DocPage.vue's scoped `.page` rule ties on
       specificity (0,2,0) and, being later in source order, would otherwise win —
       leaving the drop shadow and 12 px page gap visible in Web layout (AC-39). */
    box-shadow: none !important;
    margin-top: 0 !important;
    margin-bottom: 0 !important;
    /* AC-40: the column is the first section's content width, not the full page
       width; --web-content-width is set on the editor root (DocumentView).
       !important overrides the per-page inline `width: <pageWidth>cm`. */
    width: var(--web-content-width, auto) !important;
  }

  /* The page margins are inline padding on .page-content; drop the horizontal
     pair in Web layout so the text fills the content-width column exactly (the
     vertical padding stays as the inter-page rhythm). */
  body:not(.is-print-preview):not(.is-exporting-pdf) .is-web-layout .page-content {
    padding-left: 0 !important;
    padding-right: 0 !important;
  }

  body:not(.is-print-preview):not(.is-exporting-pdf) .is-web-layout .page-header-band,
  body:not(.is-print-preview):not(.is-exporting-pdf) .is-web-layout .page-footer-band {
    /* !important for the same specificity tie as `.page` above: DocPage.vue's
       scoped band rule sets `display: flex` and would otherwise win on source
       order, leaving the header/footer bands visible in Web layout (AC-39). */
    display: none !important;
  }
}

/* SPEC-063 AC-36: when the zoom leaves room, the paginated pages tile side by
   side in a centred grid of `--page-grid-cols` columns (set on .document-content
   by DocumentView). The grid lives inside the scaled `.pages`, so the 24px
   screen gap is written in unscaled units (÷ zoom) and each cell is one page
   wide (max-content). Only engaged above one column — a single column keeps the
   untouched normal-flow centring. The page's own auto/12px margins are
   neutralised so the grid gap alone spaces them (overlays measure live off
   `.pages`, so they follow either way). */
.document-content.is-page-grid {
  display: grid;
  grid-template-columns: repeat(var(--page-grid-cols, 1), max-content);
  justify-content: center;
  gap: calc(24px / var(--zoom, 1));
}

.document-content.is-page-grid .page {
  margin: 0;
}

/* SPEC-063 AC-52: print, print preview and PDF export always render at 100%
   Print layout, regardless of the on-screen zoom or view mode. The zoom scale
   transform is neutralised and its sizing wrapper reset to natural, and Web
   layout's hidden page chrome is restored, so the sheets and the extracted
   PDF geometry are the untransformed print layout. */
body.is-print-preview .pages-viewport--scaled,
body.is-exporting-pdf .pages-viewport--scaled {
  width: auto;
  height: auto;
}

body.is-print-preview .pages--scaled,
body.is-exporting-pdf .pages--scaled {
  position: relative;
  left: auto;
  margin-left: 0;
  transform: none;
}

body.is-print-preview .page-header-band,
body.is-print-preview .page-footer-band,
body.is-exporting-pdf .page-header-band,
body.is-exporting-pdf .page-footer-band {
  /* !important to beat `.is-web-layout .page-header-band { display: none
     !important }` above, so print, print preview and PDF export always restore
     the header/footer bands even when Web layout is on (AC-52). */
  display: revert !important;
}

@media print {
  .pages-viewport--scaled {
    width: auto !important;
    height: auto !important;
  }

  .pages--scaled {
    position: relative !important;
    left: auto !important;
    margin-left: 0 !important;
    transform: none !important;
  }

  .is-web-layout .page-header-band,
  .is-web-layout .page-footer-band {
    display: revert !important;
  }
}
</style>
