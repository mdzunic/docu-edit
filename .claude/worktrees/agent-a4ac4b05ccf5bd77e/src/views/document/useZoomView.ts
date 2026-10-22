// The document editor's zoom state (SPEC-063). Zoom is a presentation-layer
// scale on the `.pages` container only — the model, pagination and stored
// geometry stay in centimetres at 96 dpi (AC-16, AC-19). This composable
// owns the live zoom percentage, the transient fit mode, per-document
// persistence, and the client<->content coordinate helpers every overlay
// and pointer path routes through so it lands where the pointer is at any
// zoom (AC-21, AC-22).
//
// The pure arithmetic lives in lib/zoom.ts; this file is the DOM/reactive
// shell around it: it measures the live page geometry for fit modes and the
// scaled-size wrapper, and reads the live `.pages` rect for the helpers.
import { computed, ref, watch, type Ref } from "vue";
import {
  clampZoom,
  clientToContentPx as clientToContentPxPure,
  contentPxToClient as contentPxToClientPure,
  DEFAULT_ZOOM,
  fitPageZoom,
  fitWidthZoom,
  gridColumns,
  parseZoomInput,
  type Point,
  stepZoomDown,
  stepZoomUp,
  ZOOM_PRESETS,
} from "../../lib/zoom";
import { usePreferences } from "../../composables/usePreferences";

export type FitMode = "none" | "fit-width" | "fit-page";

export interface ZoomViewDeps {
  /** The open document's id — the persistence key's suffix. */
  documentId: Readonly<Ref<string>>;
  /** The `.pages` container; its live rect drives the coordinate helpers. */
  pageRef: Ref<Element | null>;
  /** The scrollable `<main>` ancestor — the fit-mode measurement box. */
  getScrollContainer: () => Element | null;
  /** Registers a view-unmount cleanup callback. */
  onUnmount: (cb: () => void) => void;
  /** True when the side-by-side page grid may engage — Print layout (incl.
   * Focus), not Web layout, whose single continuous column never tiles
   * (SPEC-063 AC-36). Defaults to always-eligible when omitted. */
  isGridEligible?: () => boolean;
}

/** The per-document persistence key for zoom (AC-29). */
export function zoomStorageKey(documentId: string): string {
  return `docu-edit:zoom:${documentId}`;
}

function readStoredZoom(documentId: string): number | null {
  if (!documentId) {
    return null;
  }
  try {
    const raw = localStorage.getItem(zoomStorageKey(documentId));
    if (raw === null) {
      return null;
    }
    const value = Number.parseInt(raw, 10);
    return Number.isNaN(value) ? null : clampZoom(value);
  } catch {
    return null;
  }
}

function writeStoredZoom(documentId: string, percent: number): void {
  if (!documentId) {
    return;
  }
  try {
    localStorage.setItem(zoomStorageKey(documentId), String(Math.round(percent)));
  } catch {
    // A private-mode / quota failure must never break editing.
  }
}

/** The first laid-out page element inside `.pages`, whose own offset
 * geometry is the 100 %-scale page size (offsetWidth/Height are unaffected
 * by an ancestor's CSS transform). Selected by the page's stable test id —
 * DocPage's root carries `data-testid="document-page"`. */
function firstPageEl(pages: Element | null): HTMLElement | null {
  if (!(pages instanceof HTMLElement)) {
    return null;
  }
  return pages.querySelector<HTMLElement>('[data-testid="document-page"]');
}

export function useZoomView(deps: ZoomViewDeps) {
  const { documentId, pageRef, getScrollContainer, onUnmount } = deps;
  const isGridEligible = deps.isGridEligible ?? (() => true);
  // The default a new document opens at comes from preferences (SPEC-066).
  const { preferences } = usePreferences();

  // The source of truth is the integer percentage; the scale fraction the
  // template and helpers use is derived from it (AC-21: EditorContext.zoom
  // is the fraction — getComputedStyle(.pages).transform reads matrix(z,…)).
  const zoomPercent = ref<number>(DEFAULT_ZOOM);
  const fitMode = ref<FitMode>("none");
  const zoom = computed(() => zoomPercent.value / 100);

  // The unscaled laid-out size of `.pages`, measured live; the sizing
  // wrapper multiplies it by the scale so scrollbars reflect the scaled
  // content (AC-17).
  const naturalWidth = ref(0);
  const naturalHeight = ref(0);

  // The scaled layout only engages when actually zoomed AND the natural page
  // size has been measured; at 100 % (the default) `.pages` keeps its
  // ordinary normal-flow centring, so nothing regresses and there is no
  // first-paint flash before the measurement lands (AC-16, AC-17).
  const scaled = computed(() => zoomPercent.value !== 100 && naturalWidth.value > 0);
  const viewportStyle = computed(() => {
    const scale = zoom.value;
    if (!scaled.value) {
      return { "--zoom": String(scale) } as Record<string, string>;
    }
    // The wrapper takes the SCALED dimensions so the scroll container's
    // scrollbars reflect the zoomed footprint (AC-17); `.pages` is then
    // centred inside it so the transform-origin:top center pivot lands on the
    // wrapper centre.
    return {
      "--zoom": String(scale),
      width: `${naturalWidth.value * scale}px`,
      height: `${naturalHeight.value * scale}px`,
    };
  });

  // The ruler sits above `.pages` as a sibling. It lays itself out at the
  // zoomed page width (every tick at its centimetre × zoom — no transform, so
  // its labels and markers never stretch) and is centred in the page column —
  // but `.pages` reserves its full scaled footprint and so left-aligns (and
  // scrolls) once that footprint overflows the column. Giving the ruler's
  // wrapper the SAME scaled-footprint layout width makes it share that
  // overflow behaviour, so ruler and page track together at every zoom
  // (AC-53/AC-54); the ruler itself is centred within this box (CSS) on the
  // footprint centre.
  const rulerSizerStyle = computed(() =>
    scaled.value
      ? ({ width: `${naturalWidth.value * zoom.value}px` } as Record<string, string>)
      : ({} as Record<string, string>),
  );

  function measureNatural(): void {
    const el = pageRef.value;
    if (el instanceof HTMLElement) {
      naturalWidth.value = el.offsetWidth;
      naturalHeight.value = el.offsetHeight;
    }
  }

  // How many pages tile side by side at the current zoom (AC-36):
  // floor(availableWidth / (pageWidth * zoom + 24)), min one. Only Print
  // layout tiles; Web layout's single column keeps this at one.
  const pageColumns = ref(1);
  function recomputeGrid(): void {
    const container = getScrollContainer();
    const page = firstPageEl(pageRef.value);
    if (!isGridEligible() || !page || !(container instanceof HTMLElement)) {
      pageColumns.value = 1;
      return;
    }
    // Never reserve more columns than there are pages: a grid wider than its
    // content parks the pages in the leading columns and leaves the trailing
    // ones empty, which pulls the (only) page off `.pages`' centre and breaks
    // the ruler/overlay registration (AC-54/AC-57). A single page therefore
    // always stays in one centred column.
    const pageCount = pageRef.value instanceof HTMLElement
      ? pageRef.value.querySelectorAll('[data-testid="document-page"]').length
      : 1;
    pageColumns.value = Math.min(
      gridColumns(container.clientWidth, page.offsetWidth, zoom.value),
      Math.max(1, pageCount),
    );
  }

  /** True when the pages should tile in a multi-column grid (AC-36) — a
   * single column keeps today's untouched normal-flow centring. */
  const pageGrid = computed(() => pageColumns.value > 1);

  /** The style bound onto the page container so the grid CSS reads the column
   * count; empty (no grid) when a single column fits. */
  const gridStyle = computed(() =>
    pageGrid.value ? ({ "--page-grid-cols": String(pageColumns.value) } as Record<string, string>) : {},
  );

  // Load the stored zoom whenever the document changes; a document with no
  // stored value opens at the "Default zoom for new documents" preference
  // (SPEC-066), which is 100 % by default (AC-30, AC-31). A fit-mode default
  // is applied by activating the fit mode; recomputeFit fills in the percent
  // once the geometry is measured.
  watch(
    documentId,
    (id) => {
      const stored = readStoredZoom(id);
      if (stored !== null) {
        zoomPercent.value = stored;
        fitMode.value = "none";
      } else if (preferences.defaultZoom === "fit-width" || preferences.defaultZoom === "fit-page") {
        fitMode.value = preferences.defaultZoom;
        recomputeFit();
      } else {
        zoomPercent.value = preferences.defaultZoom;
        fitMode.value = "none";
      }
    },
    { immediate: true },
  );

  function persist(): void {
    writeStoredZoom(documentId.value, zoomPercent.value);
  }

  /** Set an explicit percentage, clearing any fit mode (AC-8, AC-9). */
  function setZoomPercent(percent: number, mode: FitMode = "none"): void {
    zoomPercent.value = clampZoom(percent);
    fitMode.value = mode;
    persist();
  }

  /** Apply the editable field's raw text; revert (no-op) on invalid input
   * so the caller's displayed value snaps back to the current zoom (AC-10). */
  function applyZoomInput(raw: string): void {
    const parsed = parseZoomInput(raw);
    if (parsed === null) {
      return;
    }
    setZoomPercent(parsed);
  }

  function zoomIn(): void {
    setZoomPercent(stepZoomUp(zoomPercent.value));
  }

  function zoomOut(): void {
    setZoomPercent(stepZoomDown(zoomPercent.value));
  }

  function zoomReset(): void {
    setZoomPercent(DEFAULT_ZOOM);
  }

  /** Recompute a fit mode against the live geometry (AC-27): called on
   * window resize and when a side pane opens or closes. A no-op unless a fit
   * mode is active. */
  function recomputeFit(): void {
    if (fitMode.value === "none") {
      return;
    }
    const page = firstPageEl(pageRef.value);
    const container = getScrollContainer();
    if (!page || !(container instanceof HTMLElement)) {
      return;
    }
    const next =
      fitMode.value === "fit-width"
        ? fitWidthZoom(container.clientWidth, page.offsetWidth)
        : fitPageZoom(container.clientHeight, page.offsetHeight);
    zoomPercent.value = next;
    persist();
  }

  function setFitWidth(): void {
    fitMode.value = "fit-width";
    recomputeFit();
  }

  function setFitPage(): void {
    fitMode.value = "fit-page";
    recomputeFit();
  }

  /** Ctrl/Cmd + wheel and trackpad pinch (which the browser also reports as
   * a wheel event with ctrlKey) zoom the document (AC-13). preventDefault
   * stops the browser's own page zoom (AC-14). */
  function onWheel(event: WheelEvent): void {
    if (!event.ctrlKey && !event.metaKey) {
      return;
    }
    event.preventDefault();
    if (event.deltaY < 0) {
      zoomIn();
    } else if (event.deltaY > 0) {
      zoomOut();
    }
  }

  function clientToContentPx(clientX: number, clientY: number): Point {
    const el = pageRef.value;
    const rect = el instanceof Element ? el.getBoundingClientRect() : { left: 0, top: 0 };
    return clientToContentPxPure(clientX, clientY, rect, zoom.value);
  }

  function contentPxToClient(contentX: number, contentY: number): Point {
    const el = pageRef.value;
    const rect = el instanceof Element ? el.getBoundingClientRect() : { left: 0, top: 0 };
    return contentPxToClientPure(contentX, contentY, rect, zoom.value);
  }

  // Track the laid-out page size (content edits, pagination) and the
  // available width (window resize, pane toggles → recompute fit; AC-27).
  const resizeObserver =
    typeof ResizeObserver === "undefined"
      ? null
      : new ResizeObserver(() => {
          measureNatural();
          recomputeFit();
          recomputeGrid();
        });
  // Zoom changes re-tile the grid (more columns fit as you zoom out); the
  // natural width then follows via measureNatural on the next observed frame.
  watch(zoom, () => recomputeGrid());
  watch(
    pageRef,
    (el) => {
      if (resizeObserver && el instanceof Element) {
        resizeObserver.observe(el);
        measureNatural();
      }
    },
    { immediate: true },
  );

  function onWindowResize(): void {
    measureNatural();
    recomputeFit();
    recomputeGrid();
  }
  if (typeof window !== "undefined") {
    window.addEventListener("resize", onWindowResize);
  }

  onUnmount(() => {
    resizeObserver?.disconnect();
    if (typeof window !== "undefined") {
      window.removeEventListener("resize", onWindowResize);
    }
  });

  // The editorCommands entries the menu bar and keyboard shortcuts dispatch
  // through (AC-33, AC-34). Both id spellings are registered: the camelCase
  // shortcut ids (view.zoomIn/Out/Reset, AC-11) and the hyphenated menu ids
  // (view.zoom-in/out/reset/fit-width/fit-page and view.zoom-<n>, AC-34).
  const commands: Array<[string, () => void]> = [
    ["view.zoomIn", zoomIn],
    ["view.zoomOut", zoomOut],
    ["view.zoomReset", zoomReset],
    ["view.zoom-in", zoomIn],
    ["view.zoom-out", zoomOut],
    ["view.zoom-reset", zoomReset],
    ["view.zoom-fit-width", setFitWidth],
    ["view.zoom-fit-page", setFitPage],
    ...ZOOM_PRESETS.map((preset): [string, () => void] => [`view.zoom-${preset}`, () => setZoomPercent(preset)]),
  ];

  return {
    commands,
    zoom: zoom as Readonly<Ref<number>>,
    zoomPercent: zoomPercent as Readonly<Ref<number>>,
    fitMode: fitMode as Readonly<Ref<FitMode>>,
    naturalWidth: naturalWidth as Readonly<Ref<number>>,
    naturalHeight: naturalHeight as Readonly<Ref<number>>,
    /** True when the scaled (absolute + sized-wrapper) layout is engaged. */
    scaled,
    /** Inline style for the `.pages-viewport` wrapper (CSS vars + scaled
     * width/height); pair with the `.pages-viewport--scaled` / `.pages--scaled`
     * classes gated on `scaled`. */
    viewportStyle,
    /** Inline width for the ruler's sizing wrapper so it shares `.pages`'
     * scaled footprint and overflow/centring behaviour (AC-53/AC-54); empty
     * at 100 %. Pair with the `.ruler-sizer` class. */
    rulerSizerStyle,
    /** True when the pages tile in a multi-column grid (AC-36). */
    pageGrid,
    /** How many columns the grid tiles into (AC-36). */
    pageColumns: pageColumns as Readonly<Ref<number>>,
    /** Style carrying `--page-grid-cols` for the grid container; empty at one
     * column. */
    gridStyle,
    /** Recompute the column count — call when the view mode toggles between
     * Print and Web layout (grid eligibility changes). */
    recomputeGrid,
    setZoomPercent,
    applyZoomInput,
    zoomIn,
    zoomOut,
    zoomReset,
    setFitWidth,
    setFitPage,
    recomputeFit,
    onWheel,
    clientToContentPx,
    contentPxToClient,
  };
}
