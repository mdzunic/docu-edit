// PDF first-page rasterisation for the library's previews, on Mozilla's
// pdf.js. This is the ONE sanctioned exception to the frozen dependency set
// (approved 2026-08-30): `pdfjs-dist` — Apache-2.0, actively maintained by
// Mozilla, and with ZERO runtime dependencies (its sole optional dependency,
// a native Node canvas, is not installed; we only render in the browser).
//
// The library is loaded lazily, inside the first rasterisation call, so the
// app bundle and its paint budgets never carry pdf.js; the worker ships as
// its own Vite chunk fetched on demand. Callers gate on
// `pdfRasterSupported()` first — a host without Workers (jsdom, older
// engines) keeps the aspect-true paper placeholder instead, and any render
// failure must degrade the same way (this module throws, it never lies).

/** Whether this host can rasterise at all: pdf.js needs a real Worker and a
 * real 2D canvas — jsdom (unit tests) has neither. */
export function pdfRasterSupported(): boolean {
  return typeof Worker === "function" && typeof document !== "undefined";
}

let workerConfigured = false;

/** Renders the first page of `bytes` to a PNG data URL whose longer edge is
 * `maxEdgePx`. Page rotation is honoured (pdf.js folds /Rotate into the
 * default viewport). Throws on any parse or render failure — the caller owns
 * the placeholder fallback. */
export async function rasterisePdfFirstPage(bytes: Uint8Array, maxEdgePx: number): Promise<string> {
  // The LEGACY build, deliberately: the modern build leans on brand-new
  // platform features (e.g. Uint8Array.prototype.toHex) that the pinned
  // Playwright 1.49 Chromium lacks — the legacy build carries the polyfills.
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  if (!workerConfigured) {
    const worker = await import("pdfjs-dist/legacy/build/pdf.worker.min.mjs?url");
    pdfjs.GlobalWorkerOptions.workerSrc = worker.default;
    workerConfigured = true;
  }
  // pdf.js transfers the buffer to its worker; hand it a copy so the model's
  // own bytes (round-trip fidelity!) can never be detached.
  //
  // Nothing here has to opt out of eval for SEC-6's `script-src 'self'`:
  // pdf.js v6 dropped the `new Function` path it used to compile font
  // programs and colour-space code through, along with the `isEvalSupported`
  // option that disabled it — neither `pdf.mjs` nor `pdf.worker.mjs` contains
  // a single `new Function(` or `eval(` any more. Worth re-checking on a
  // major pdf.js upgrade: if eval comes back, this call needs the opt-out
  // again rather than the CSP being widened to admit it.
  const task = pdfjs.getDocument({ data: bytes.slice() });
  try {
    const doc = await task.promise;
    const page = await doc.getPage(1);
    const base = page.getViewport({ scale: 1 });
    const scale = maxEdgePx / Math.max(base.width, base.height, 1);
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(viewport.width));
    canvas.height = Math.max(1, Math.round(viewport.height));
    // `intent: "print"`: the display intent schedules its paint loop on
    // requestAnimationFrame, which never fires in a hidden/non-compositing
    // tab — the render would hang forever exactly where thumbnails build
    // (background tabs, embedded panes). The print intent continues without
    // rAF and skips interactive annotations: right for a static raster.
    await page.render({ canvas, viewport, intent: "print" }).promise;
    return canvas.toDataURL("image/png");
  } finally {
    await task.destroy();
  }
}
