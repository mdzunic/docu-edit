// The print preview and print flow (SPEC-034 AC-1..6, AC-13, AC-27..31).
//
// The preview IS the paginated canvas: opening it overlays the range panel
// and marks the body with `is-print-preview`, whose global rules (App.vue)
// hide every screen-only affordance — the same set the printed output
// excludes — and hide pages outside the selected range. Printing injects an
// @page rule sized to the document's own page geometry with zero printer
// margin (the rendered page boxes carry the document margins themselves),
// so the browser's sheets match the on-screen pages exactly and the
// preview's pagination is, by construction, the printed pagination.
//
// Virtualized pages render as empty boxes, so before printing (or PDF
// export) every page must realize: `forceRealizeAll` switches the view out
// of virtualization and `realizeAllPages()` waits until the layout settles
// with no virtual page boxes left, bounded by a deadline well under the
// export budget. Dismissing the preview mid-realize aborts the wait.
import { computed, nextTick, onScopeDispose, ref, watchEffect } from "vue";
import { isDocumentContentEmpty, parsePageRange, printPageRuleCss } from "../../lib/printRange";
import type { PaginationController } from "../../composables/usePagination";
import type { EditorContext } from "./context";

export interface PrintPreviewDeps extends EditorContext {
  pagination: PaginationController;
  /** SPEC-070: re-routes the connectors once every page is realized.
   * Connectors are measured off the DOM, and a virtualized page has none of
   * its shapes in it — a connector attached to one resolved to its frozen
   * free position until the page materialized. Both callers of
   * `realizeAllPages` (this preview and the PDF export) need the re-route,
   * so it happens here rather than in each of them. */
  remeasureConnectors?: () => void;
}

/** How long realizeAllPages may wait for every page to render and measure
 * before giving up — well inside the 10s export budget (AC-21). */
const REALIZE_DEADLINE_MS = 8000;

function raf(): Promise<void> {
  return new Promise((resolve) => {
    if (typeof requestAnimationFrame === "function") {
      requestAnimationFrame(() => resolve());
    } else {
      setTimeout(resolve, 16);
    }
  });
}

export function usePrintPreview(deps: PrintPreviewDeps) {
  const printPreviewOpen = ref(false);
  /** True from just before window.print() until it returns — the template
   * drops the formatting-marks class off this (AC-13). */
  const printing = ref(false);
  /** Disables page virtualization so every page renders its content. */
  const forceRealizeAll = ref(false);

  const rangeMode = ref<"all" | "custom">("all");
  const rangeText = ref("");

  /** Pages the preview/print can address: zero for a content-empty
   * document, which previews as no pages and cannot print (AC-27). */
  const printablePageCount = computed(() => {
    const model = deps.doc.value;
    if (!model || isDocumentContentEmpty(model)) {
      return 0;
    }
    return deps.pagination.pages.value.length;
  });

  const rangeResult = computed(() =>
    rangeMode.value === "custom"
      ? parsePageRange(rangeText.value, printablePageCount.value)
      : { pages: undefined, error: undefined },
  );

  const rangeError = computed(() =>
    printablePageCount.value > 0 ? rangeResult.value.error : undefined,
  );

  const canPrint = computed(() => printablePageCount.value > 0 && rangeError.value === undefined);

  /** 0-based indexes of laid-out pages the preview hides and printing
   * skips (AC-29/31): everything for an empty document, the complement of
   * a valid custom range otherwise. */
  const printHiddenPages = computed<Set<number>>(() => {
    const total = deps.pagination.pages.value.length;
    const hidden = new Set<number>();
    if (printablePageCount.value === 0) {
      for (let index = 0; index < total; index += 1) {
        hidden.add(index);
      }
      return hidden;
    }
    const selected = rangeMode.value === "custom" ? rangeResult.value.pages : undefined;
    if (selected) {
      const wanted = new Set(selected);
      for (let index = 0; index < total; index += 1) {
        if (!wanted.has(index + 1)) {
          hidden.add(index);
        }
      }
    }
    return hidden;
  });

  function openPrintPreview(): void {
    printPreviewOpen.value = true;
    forceRealizeAll.value = true;
    void realizeAllPages();
  }

  function closePrintPreview(): void {
    printPreviewOpen.value = false;
    forceRealizeAll.value = false;
  }

  /** Realizes every page and waits for pagination to settle. Resolves false
   * when dismissed mid-wait or the deadline passes. */
  async function realizeAllPages(): Promise<boolean> {
    forceRealizeAll.value = true;
    await nextTick();
    deps.pagination.schedule();
    const deadline = performance.now() + REALIZE_DEADLINE_MS;
    for (;;) {
      if (!forceRealizeAll.value) {
        return false; // dismissed while realizing — stop waiting
      }
      const root = deps.pageRef.value;
      if (
        root &&
        deps.pagination.state.value === "clean" &&
        root.querySelector(".page[data-virtual]") === null
      ) {
        deps.remeasureConnectors?.();
        // One more frame so the re-routed paint layers are in the DOM before
        // the caller reads (or prints) the pages.
        await raf();
        return true;
      }
      if (performance.now() > deadline) {
        return false;
      }
      await raf();
    }
  }

  /** The preview's Print button (AC-5): realize everything, then hand off
   * to the browser's native print dialog. */
  async function confirmPrint(): Promise<void> {
    if (!canPrint.value || printing.value) {
      return;
    }
    await realizeAllPages();
    printing.value = true;
    await nextTick();
    await raf();
    try {
      window.print();
    } finally {
      printing.value = false;
    }
  }

  // The @page rule (AC-6): kept in <head> and synced to the document's own
  // page size for as long as the view lives, so even a print started from
  // the browser menu uses the document geometry.
  let pageStyleEl: HTMLStyleElement | null = null;
  watchEffect(() => {
    const model = deps.doc.value;
    if (typeof document === "undefined") {
      return;
    }
    if (!model) {
      pageStyleEl?.remove();
      pageStyleEl = null;
      return;
    }
    if (!pageStyleEl) {
      pageStyleEl = document.createElement("style");
      pageStyleEl.setAttribute("data-testid", "print-page-style");
      document.head.appendChild(pageStyleEl);
    }
    pageStyleEl.textContent = printPageRuleCss(model);
  });

  // The body marker class scoping the preview's screen-only exclusions.
  watchEffect(() => {
    if (typeof document !== "undefined") {
      document.body.classList.toggle("is-print-preview", printPreviewOpen.value);
    }
  });

  onScopeDispose(() => {
    pageStyleEl?.remove();
    pageStyleEl = null;
    if (typeof document !== "undefined") {
      document.body.classList.remove("is-print-preview");
    }
  });

  return {
    printPreviewOpen,
    printing,
    forceRealizeAll,
    rangeMode,
    rangeText,
    rangeError,
    canPrint,
    printablePageCount,
    printHiddenPages,
    openPrintPreview,
    closePrintPreview,
    confirmPrint,
    realizeAllPages,
  };
}
