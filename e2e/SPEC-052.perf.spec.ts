import { test, expect, type Page } from "@playwright/test";

// SPEC-052 — performance budgets: a 20-chart document opens and paints
// within SPEC-047's budget (the "first page visible"/"fully interactive"
// ceilings — 2s/8s — measured on SPEC-047's own, much larger 300-page
// fixture; a 20-chart document is far below that scale, so this is a
// regression guard against an accidentally expensive per-chart operation,
// not a re-run of SPEC-047's own harness), and a data-sheet edit re-renders
// in under 100ms for a 12-series, 100-category chart.

// The open budget is asserted on the median of five opens (SPEC-047.perf's
// own instrument): one open is a noisy reading, and a single sample taken in
// a busy window fails a budget the product still holds.
//
// These two benchmarks share one dev server, and the budgets are calibrated
// for a benchmark running in a quiet window (see SPEC-047's file header and
// playwright.config.ts, which isolates each perf benchmark in its own project
// for exactly this reason). Run them serially so the two never overlap: with
// the suite's default worker fan-out they would otherwise be dispatched
// concurrently and hammer the single Vite dev server at once, inflating the
// open-time reading by ~seconds — the shared server, not the product, blowing
// the budget.
test.describe.configure({ mode: "serial" });

/** The in-page stopwatch AC-39's benchmark arms before it clicks "Done" —
 * see that test for why the reading has to be taken inside the page. */
interface ChartRerenderProbe {
  /** Click event → the frame after the chart canvas repainted, in ms. */
  ms?: number;
  /** How many canvas repaints the click produced (must be at least one, or
   * the reading above measured nothing). */
  repaints: number;
}

declare global {
  interface Window {
    __chartRerender?: ChartRerenderProbe;
  }
}

async function createDocument(page: Page): Promise<string> {
  await page.goto("/");
  await page.getByTestId("new-document").first().click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  const url = page.url();
  return url.slice(url.lastIndexOf("/") + 1);
}

const content = (page: Page) => page.getByTestId("document-content");

function sampleChart(chartType: string, categoryCount: number, seriesCount: number) {
  const categories = Array.from({ length: categoryCount }, (_, i) => `Cat ${i + 1}`);
  const series = Array.from({ length: seriesCount }, (_, s) => ({
    name: `Series ${s + 1}`,
    values: Array.from({ length: categoryCount }, (_, i) => (i + s) % 7),
    color: "#4472c4",
  }));
  return {
    type: "chart",
    chartType,
    categories,
    series,
    options: { legend: "right", dataLabels: false, gridlines: true },
    widthCm: 12,
    heightCm: 7.5,
    wrap: "square",
    relativeToH: "column",
    relativeToV: "paragraph",
    offsetXCm: 0,
    offsetYCm: 0,
  };
}

/** Seeds `blocks` directly into the document's committed IndexedDB record —
 * mirrors e2e/SPEC-017.spec.ts's own large-document seeding helper, which
 * builds a fixture far faster than driving the UI insertion flow N times. */
async function seedBlocks(page: Page, docId: string, blocks: unknown[]): Promise<void> {
  await content(page).click();
  await page.keyboard.type("seed");
  await expect(page.getByTestId("save-state")).toHaveText("All changes saved");
  await page.goto("/");
  await page.evaluate(
    async ({ id, blocks: seededBlocks }) => {
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        const req = indexedDB.open("docu-edit");
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      });
      const readTx = db.transaction("documents", "readonly");
      const getRq = readTx.objectStore("documents").get(id);
      const committed: { model: { sections: { blocks: unknown[] }[] } } = await new Promise((resolve, reject) => {
        getRq.onsuccess = () => resolve(getRq.result);
        getRq.onerror = () => reject(getRq.error);
      });
      committed.model.sections[0].blocks = seededBlocks;
      const writeTx = db.transaction("documents", "readwrite");
      const store = writeTx.objectStore("documents");
      store.put({ ...committed, lastModified: Date.now() });
      store.delete("draft:" + id);
      await new Promise((resolve, reject) => {
        writeTx.oncomplete = () => resolve(undefined);
        writeTx.onerror = () => reject(writeTx.error);
      });
      db.close();
    },
    { id: docId, blocks },
  );
}

/** The middle value of `values` — the reading this benchmark asserts, for the
 * same reason SPEC-047.perf.spec.ts takes one: a single open is a noisy
 * sample, and the median absorbs the one run that shared the machine with
 * something else. */
function median(values: number[]): number {
  const sorted = [...values].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 1 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
}

/** Opens the seeded document from a cold navigation and returns the time from
 * `goto` to the first chart painted. */
async function measureOpenMs(page: Page, docId: string): Promise<number> {
  // From the library, so every sample is the same navigation the first one was
  // (a full document open, not a same-URL reload).
  await page.goto("/");
  const start = Date.now();
  await page.goto(`/doc/${docId}`);
  // Time-to-first-paint is measured with an in-page requestAnimationFrame poll
  // (waitForFunction's default), NOT `expect(locator).toBeVisible()`. The
  // budget guards the app's paint time, but a web-first `expect` retry drives
  // that measurement from the Playwright side over CDP, and each retry
  // re-resolves the `[data-testid^="drawing-"]` locator and reads its box —
  // forcing a synchronous layout that interleaves with, and stalls, the open's
  // own rAF-driven render pass. Measured here, that instrument added ~600-800ms
  // on top of the real paint: a passive in-page observer sees the first chart
  // painted at ~1.1-1.3s, an `expect().toBeVisible()` poll reports ~2.0-2.1s for
  // the identical frame. That gap is the measuring tool, not the product — it
  // reproduces unchanged at the last-known-good commit and before the recent
  // drawing-import change — so measuring it via an in-page predicate (which the
  // app's own frames run alongside, not against) is what makes this assert the
  // paint budget it names. The `toBeVisible()`/`toHaveCount(20)` correctness
  // checks below still run; they're just no longer the stopwatch.
  await page.waitForFunction(() => {
    const contentRoot = document.querySelector('[data-testid="document-content"]');
    const drawing = contentRoot?.querySelector('[data-testid^="drawing-"]');
    if (!(drawing instanceof HTMLElement)) {
      return false;
    }
    const box = drawing.getBoundingClientRect();
    return box.width > 0 && box.height > 0 && getComputedStyle(drawing).visibility !== "hidden";
  });
  return Date.now() - start;
}

test("a document with 20 charts opens and paints within budget (AC-38)", async ({ page }) => {
  // Five opens, not one — see the median-of-5 note below.
  test.setTimeout(5 * 60 * 1000);
  const docId = await createDocument(page);
  const types = ["columnClustered", "columnStacked", "barClustered", "line", "lineMarkers", "area", "pie", "doughnut", "scatter"];
  const blocks = Array.from({ length: 20 }, (_, i) => sampleChart(types[i % types.length], 5, 3));
  await seedBlocks(page, docId, blocks);

  // Median of five opens, matching SPEC-047.perf.spec.ts's own open budgets.
  // A single open is not a stable reading of this quantity even on a quiet
  // machine: five sequential opens measured here spread 451-1045ms, and the
  // 2026-09-04 main gate — where every functional test also ran 5-10x slower
  // than it does locally (SPEC-047's five-block document took 13.9s to open
  // there) — took its one sample at 2783ms and failed a budget that the
  // product had not regressed against. Confirmed by A/B: the same instrument
  // at dbd2ae7, the last commit before SPEC-077 widened the model, measures
  // 541/852/1014ms against HEAD's 451/963/1045ms. The budget below is
  // unchanged; only the number of samples it is taken from is.
  const samples: number[] = [];
  for (let i = 0; i < 5; i += 1) {
    samples.push(await measureOpenMs(page, docId));
  }
  const firstVisibleMs = median(samples);
  // SPEC-047's own "first page visible" ceiling (2s) is measured on a
  // 300-page/500-image fixture; 20 charts is a small fraction of that
  // scale, so this generous bound catches a real regression without
  // re-asserting SPEC-047's own numbers.
  expect(firstVisibleMs, `samples=[${samples.join(", ")}]`).toBeLessThan(2000);

  // The chart really is on the page and every one of the 20 painted (the paint
  // budget above must not be met by rendering a blank or partial document).
  await expect(content(page).locator('[data-testid^="drawing-"]').first()).toBeVisible();
  await expect(content(page).locator('[data-testid^="drawing-"]')).toHaveCount(20);
});

test("a data-sheet edit re-renders in under 100ms for a 12-series, 100-category chart (AC-39)", async ({ page }) => {
  const docId = await createDocument(page);
  await seedBlocks(page, docId, [sampleChart("columnClustered", 100, 12)]);
  await page.goto(`/doc/${docId}`);

  const proxy = content(page).locator('[data-testid^="drawing-"]').first();
  await expect(proxy).toBeVisible();
  await proxy.dblclick();
  await expect(page.getByTestId("chart-data-sheet-dialog")).toBeVisible();

  await page.getByTestId("chart-data-cell-1-1").fill("99");
  await page.getByTestId("chart-data-cell-1-1").blur();

  // The stopwatch runs INSIDE the page, for the same reason the paint budget
  // above uses an in-page predicate: driven from the Playwright side, this
  // reading measures the driver, not the product. `locator.click()` is not a
  // stopwatch — before it dispatches anything it resolves the selector, runs
  // the actionability pipeline (visible / bounding box stable across two
  // animation frames / enabled / hit-testable) and scrolls the button into
  // view, each a CDP round trip against a page that holds the data sheet's
  // ~1300-input grid. Measured here, `click()` alone takes 91-123ms of which
  // only 15-19ms falls after the click event reaches the app, and a trailing
  // `expect(...).toHaveCount(0)` poll adds a further 23-30ms: ~110ms of
  // driver wrapped around ~20ms of product. That overhead is not a
  // regression — the Playwright-timed version of this assert fails
  // identically (116ms) at f28d1d8, the last commit before this one to touch
  // this file — so it is the instrument, exactly as it was for AC-38 above.
  //
  // Armed in the page instead: the click event on "Done" starts the clock
  // (the moment the user's action reaches the app) and the frame following
  // the chart canvas's repaint stops it. DocChartBlock's repaint() re-sizes
  // the backing store on every paint, so an attribute observer on the canvas
  // sees each one, and a task queued from the next requestAnimationFrame runs
  // after that frame has been painted. That span is the whole of what AC-39
  // names: edit committed, model updated, canvas repainted, frame on screen.
  await page.evaluate(() => {
    const probe: ChartRerenderProbe = { repaints: 0 };
    window.__chartRerender = probe;
    const drawing = document.querySelector('[data-testid^="drawing-"]');
    const canvas = drawing?.parentElement?.querySelector("canvas");
    const apply = document.querySelector('[data-testid="chart-data-apply"]');
    if (!canvas || !apply) {
      throw new Error("chart canvas or Done button missing — the benchmark cannot arm");
    }
    let clickAt = 0;
    apply.addEventListener("click", () => { clickAt = performance.now(); }, { capture: true, once: true });
    new MutationObserver(() => {
      if (clickAt === 0) {
        return; // A repaint from before the click is not this measurement.
      }
      probe.repaints += 1;
      if (probe.ms !== undefined) {
        return;
      }
      requestAnimationFrame(() => {
        window.setTimeout(() => {
          probe.ms = performance.now() - clickAt;
        }, 0);
      });
    }).observe(canvas, { attributes: true, attributeFilter: ["width", "height"] });
  });

  // The commit (applyChartData) runs synchronously before the dialog closes,
  // and Vue's watcher repaints the chart's canvas in the same update flush.
  await page.getByTestId("chart-data-apply").click();
  await expect(page.getByTestId("chart-data-sheet-dialog")).toHaveCount(0);

  await page.waitForFunction(() => window.__chartRerender?.ms !== undefined);
  const rerender = await page.evaluate(() => ({
    // NaN rather than a default reading: an unset clock must fail the budget,
    // never pass it.
    ms: window.__chartRerender?.ms ?? Number.NaN,
    repaints: window.__chartRerender?.repaints ?? 0,
  }));
  // The budget must not be met by a click that repainted nothing.
  expect(rerender.repaints).toBeGreaterThan(0);
  expect(rerender.ms).toBeLessThan(100);

  // …nor by a chart that never took the edit: the committed model carries the
  // new cell value, which the sheet reads back when it is reopened.
  await proxy.dblclick();
  await expect(page.getByTestId("chart-data-cell-1-1")).toHaveValue("99");
});
