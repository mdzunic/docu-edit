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
}

html,
body,
#app {
  height: 100%;
  margin: 0;
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

  /* App chrome never prints: only the page boxes reach the sheets. */
  .app-bar,
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
}
</style>
