import { test, expect, type Page, type Browser } from "@playwright/test";
import { existsSync, readFileSync } from "node:fs";
import { execFileSync } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { MOD } from "./keys";
import { SPEC137_LARGE_XLSX } from "../src/lib/spreadsheet/fixtures/build";

// SPEC-137 — large workbooks: the benchmark. Every row of specs/137's "The
// budgets" table, asserted on the generated fixture
// (scripts/build-spec137-fixtures.ts) under SPEC-047's two-gate rule
// (absolute ceiling; the checked-in e2e/perf-baseline.json value + 15%),
// median of 5 — exactly the method e2e/SPEC-047.perf.spec.ts established.
//
// Timing validity (restating SPEC-047's own notes, which apply here
// unchanged): the baselines below are calibrated for THIS FILE running with
// the machine otherwise quiet. playwright.config.ts dispatches it in its own
// project (`perf-workbook`), chained one level past `perf-large-document` —
// today's chain tail — so it runs alone, after the whole functional suite
// and every other benchmark have drained. A timing failure seen only in a
// full-suite or whole-gate run should be confirmed with a SOLO run of this
// file before being treated as a regression: what overlaps the start of a
// full run (the last in-flight functional tests, a sibling benchmark's
// tail) lands inside the first of the five runs, which the median absorbs,
// but host-level contention does not respect that — the factory's gate
// container shares its host machine's CPU, so heavy work on the host during
// a gate run (builds, other lanes, review agents) starves the container and
// inflates every row together. A gate failure whose rows all drifted in
// unison is contention, not a regression; a quiet re-run is the diagnostic.
const REPO_ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), "..");
const LARGE_FIXTURE_PATH = path.join(REPO_ROOT, "e2e", "fixtures", "spec137-large.xlsx");
const LIBRARY_FIXTURE_PATH = path.join(REPO_ROOT, "e2e", "fixtures", "spec137-library.xlsx");
const BASELINE_PATH = path.join(REPO_ROOT, "e2e", "perf-baseline.json");
const RUNS = 5;

// --- "The budgets" table, restated as data (AC-2) --------------------------
const BUDGETS: Record<string, { ceiling: number; direction: "atMost" | "atLeast" }> = {
  sheetOpenFirstPaintMs: { ceiling: 6000, direction: "atMost" },
  sheetMainThreadBlockDuringOpenMs: { ceiling: 150, direction: "atMost" },
  sheetReopenFirstPaintMs: { ceiling: 3000, direction: "atMost" },
  sheetScrollFps: { ceiling: 30, direction: "atLeast" },
  sheetMainThreadBlockDuringScrollMs: { ceiling: 50, direction: "atMost" },
  sheetJumpToEndMs: { ceiling: 500, direction: "atMost" },
  sheetChainRecalcMs: { ceiling: 2000, direction: "atMost" },
  sheetMainThreadBlockDuringRecalcMs: { ceiling: 50, direction: "atMost" },
  sheetSaveMs: { ceiling: 5000, direction: "atMost" },
  sheetMainThreadBlockDuringSaveMs: { ceiling: 100, direction: "atMost" },
  sheetLibraryLandingMs: { ceiling: 1500, direction: "atMost" },
  sheetLibraryHeapDeltaMB: { ceiling: 15, direction: "atMost" },
  sheetHeapAfterSettleMB: { ceiling: 300, direction: "atMost" },
};

function buildFixturesIfMissing(): void {
  if (existsSync(LARGE_FIXTURE_PATH) && existsSync(LIBRARY_FIXTURE_PATH)) {
    return;
  }
  execFileSync("npx", ["vite-node", "scripts/build-spec137-fixtures.ts"], { cwd: REPO_ROOT, stdio: "inherit" });
  if (!existsSync(LARGE_FIXTURE_PATH) || !existsSync(LIBRARY_FIXTURE_PATH)) {
    throw new Error("SPEC-137 fixture build did not produce both e2e/fixtures/spec137-*.xlsx files");
  }
}

interface Baseline {
  _comment?: string;
  [key: string]: number | string | undefined;
}

function readBaseline(): Baseline {
  if (!existsSync(BASELINE_PATH)) {
    return {};
  }
  return JSON.parse(readFileSync(BASELINE_PATH, "utf-8")) as Baseline;
}

function median(values: number[]): number {
  const sorted = [...values].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];
}

/** Both of SPEC-047's gates (absolute ceiling, baseline + 15%), evaluated on
 * the median of `samples`. `expect.soft` so one budget's failure never hides
 * another's in the same run. */
function assertBudget(name: string, samples: number[], baseline: Baseline): void {
  const budget = BUDGETS[name];
  const value = median(samples);
  const baselineValue = typeof baseline[name] === "number" ? (baseline[name] as number) : undefined;
  // eslint-disable-next-line no-console
  console.log(
    `[SPEC-137] ${name}: samples=[${samples.map((n) => n.toFixed(1)).join(", ")}] median=${value.toFixed(2)} ` +
      `ceiling=${budget.direction === "atMost" ? "<=" : ">="}${budget.ceiling} baseline=${baselineValue ?? "none"}`,
  );
  if (budget.direction === "atMost") {
    expect.soft(value, `${name} absolute gate (ceiling ${budget.ceiling})`).toBeLessThanOrEqual(budget.ceiling);
    if (baselineValue !== undefined) {
      expect.soft(value, `${name} relative gate (baseline ${baselineValue} + 15%)`).toBeLessThanOrEqual(baselineValue * 1.15);
    }
  } else {
    expect.soft(value, `${name} absolute gate (floor ${budget.ceiling})`).toBeGreaterThanOrEqual(budget.ceiling);
    if (baselineValue !== undefined) {
      expect.soft(value, `${name} relative gate (baseline ${baselineValue} - 15%)`).toBeGreaterThanOrEqual(baselineValue * 0.85);
    }
  }
}

// --- Long Tasks API collector (SPEC-047's pattern, duplicated per-file —
// no shared e2e helpers module exists in this repo). ------------------------
async function installLongTaskCollector(page: Page): Promise<void> {
  await page.addInitScript(() => {
    const w = window as unknown as { __longTasks: number[]; __longTasksSupported: boolean };
    w.__longTasks = [];
    w.__longTasksSupported = false;
    try {
      const observer = new PerformanceObserver((list) => {
        for (const entry of list.getEntries()) {
          (window as unknown as { __longTasks: number[] }).__longTasks.push(entry.duration);
        }
      });
      observer.observe({ entryTypes: ["longtask"] });
      w.__longTasksSupported = true;
    } catch {
      // Left false; the budget test fails loudly on it.
    }
  });
}

async function longTasksSupported(page: Page): Promise<boolean> {
  return page.evaluate(() => (window as unknown as { __longTasksSupported?: boolean }).__longTasksSupported === true);
}

async function resetLongTasks(page: Page): Promise<void> {
  await page.evaluate(() => {
    (window as unknown as { __longTasks: number[] }).__longTasks = [];
  });
}

async function longestTask(page: Page): Promise<number> {
  const tasks = await page.evaluate(() => (window as unknown as { __longTasks: number[] }).__longTasks ?? []);
  return tasks.length > 0 ? Math.max(...tasks) : 0;
}

/** CDP `Performance.getMetrics` -> `JSHeapUsedSize`, after
 * `HeapProfiler.collectGarbage` (an unforced sample measures garbage that
 * has not been collected yet, not memory actually held). */
async function measureHeapMB(page: Page): Promise<number> {
  const client = await page.context().newCDPSession(page);
  await client.send("Performance.enable");
  await client.send("HeapProfiler.enable");
  await client.send("HeapProfiler.collectGarbage");
  const { metrics } = await client.send("Performance.getMetrics");
  await client.detach();
  const bytes = metrics.find((m) => m.name === "JSHeapUsedSize")?.value ?? 0;
  return bytes / (1024 * 1024);
}

/** The first `.sheet-cell` with non-empty text inside the grid is in the DOM
 * — the open/reopen "first paint" signal every row below builds on. */
async function waitForFirstPaintedCell(page: Page, timeout = 30_000): Promise<void> {
  await page.waitForFunction(
    () => {
      const grid = document.querySelector('[data-testid="spreadsheet-grid"]');
      if (!grid) {
        return false;
      }
      for (const cell of grid.querySelectorAll(".sheet-cell")) {
        if ((cell.textContent ?? "").trim().length > 0) {
          return true;
        }
      }
      return false;
    },
    { timeout },
  );
}

function numberFromCellText(text: string): number {
  return Number.parseFloat(text.replace(/,/g, ""));
}

interface OpenSample {
  sheetOpenFirstPaintMs: number;
  sheetMainThreadBlockDuringOpenMs: number;
}

/** t0 once open-document-input is attached (the barrier SPEC-047 found:
 * `page.goto` resolving is not the app being mounted); setInputFiles -> the
 * grid's first painted cell; the open block is the longest task from t0
 * through two seconds past that paint (initial persistence included). */
async function measureOpen(page: Page): Promise<OpenSample> {
  await page.goto("/");
  await resetLongTasks(page);
  await page.getByTestId("open-document-input").waitFor({ state: "attached" });
  const navigated = page.waitForURL(/\/sheets\/.+/, { timeout: 30_000 });
  const t0 = Date.now();
  await page.getByTestId("open-document-input").setInputFiles(LARGE_FIXTURE_PATH);
  await navigated;
  await waitForFirstPaintedCell(page);
  const sheetOpenFirstPaintMs = Date.now() - t0;
  await page.waitForTimeout(2000);
  const sheetMainThreadBlockDuringOpenMs = await longestTask(page);
  return { sheetOpenFirstPaintMs, sheetMainThreadBlockDuringOpenMs };
}

/** After a reload (the workbook not resident): close back to the library,
 * reload, then click its library-row -> first painted cell. */
async function measureReopen(page: Page): Promise<{ sheetReopenFirstPaintMs: number }> {
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();
  await page.reload();
  await expect(page.getByTestId("library-row").first()).toBeVisible();
  const t0 = Date.now();
  await page.getByTestId("library-row").first().dblclick();
  await expect(page).toHaveURL(/\/sheets\/.+/);
  await waitForFirstPaintedCell(page);
  return { sheetReopenFirstPaintMs: Date.now() - t0 };
}

/** A 5-second programmatic fling of the grid's own scroller (SPEC-116's
 * million-row test confirms `[data-testid=spreadsheet-grid]` IS the
 * scrollable element) — 2,000px steps down for the first half, up for the
 * second, sampling rAF timestamps for fps. */
async function measureScroll(page: Page): Promise<{ sheetScrollFps: number; sheetMainThreadBlockDuringScrollMs: number }> {
  await resetLongTasks(page);
  const sheetScrollFps = await page.evaluate(() => {
    const el = document.querySelector('[data-testid="spreadsheet-grid"]') as HTMLElement | null;
    if (!el) {
      return 0;
    }
    const durationMs = 5000;
    const stepPx = 2000;
    const maxScroll = Math.max(0, el.scrollHeight - el.clientHeight);
    return new Promise<number>((resolve) => {
      const start = performance.now();
      let frames = 0;
      function step() {
        frames++;
        const elapsed = performance.now() - start;
        if (elapsed < durationMs / 2) {
          el!.scrollTop = Math.min(maxScroll, el!.scrollTop + stepPx);
        } else {
          el!.scrollTop = Math.max(0, el!.scrollTop - stepPx);
        }
        if (elapsed < durationMs) {
          requestAnimationFrame(step);
        } else {
          resolve(frames / (elapsed / 1000));
        }
      }
      requestAnimationFrame(step);
    });
  });
  const sheetMainThreadBlockDuringScrollMs = await longestTask(page);
  return { sheetScrollFps, sheetMainThreadBlockDuringScrollMs };
}

/** Ctrl/Cmd+End keydown -> `[data-ref=T25000]` attached AND the name box
 * reads "T25000" (SPEC-117's "Ctrl/Cmd+End = last used cell"). */
async function measureJumpToEnd(page: Page): Promise<{ sheetJumpToEndMs: number }> {
  await page.locator('[data-testid="sheet-cell"][data-ref="A1"]').first().click();
  const t0 = Date.now();
  await page.keyboard.press(`${MOD}+End`);
  await expect(page.locator('[data-testid="sheet-cell"][data-ref="T25000"]')).toBeAttached({ timeout: 10_000 });
  await expect(page.getByTestId("name-box")).toHaveValue("T25000", { timeout: 10_000 });
  return { sheetJumpToEndMs: Date.now() - t0 };
}

/** Type a new number into S1 + Enter -> `data-calc-state` idle (first making
 * sure the pass actually started, so a vacuously-already-idle read can't
 * pass this trivially), then — untimed — T10000 holds the new chain total;
 * the longest task in the timed window. */
async function measureChainRecalc(page: Page): Promise<{ sheetChainRecalcMs: number; sheetMainThreadBlockDuringRecalcMs: number; newS1: number }> {
  const grid = page.getByTestId("spreadsheet-grid");
  await page.getByTestId("name-box").fill("S1");
  await page.keyboard.press("Enter");
  const t1Before = await page.locator('[data-testid="sheet-cell"][data-ref="T1"]').textContent().catch(() => null);
  await resetLongTasks(page);
  const newS1 = 9999.99;
  const t0 = Date.now();
  await page.keyboard.type(String(newS1));
  await page.keyboard.press("Enter");
  await Promise.race([
    expect(grid).toHaveAttribute("data-calc-state", "calculating", { timeout: 5_000 }).catch(() => undefined),
    page
      .waitForFunction(
        (previous) => document.querySelector('[data-ref="T1"]')?.textContent !== previous,
        t1Before,
        { timeout: 5_000 },
      )
      .catch(() => undefined),
  ]);
  await expect(grid).toHaveAttribute("data-calc-state", "idle", { timeout: 30_000 });
  const sheetChainRecalcMs = Date.now() - t0;
  const sheetMainThreadBlockDuringRecalcMs = await longestTask(page);
  return { sheetChainRecalcMs, sheetMainThreadBlockDuringRecalcMs, newS1 };
}

/** Untimed: navigates to T10000 and reads its numeric text against the
 * fixture's chain facts shifted by the edit just made to S1. */
async function assertChainTotal(page: Page, newS1: number): Promise<void> {
  await page.getByTestId("name-box").fill("T10000");
  await page.keyboard.press("Enter");
  const cell = page.locator('[data-testid="sheet-cell"][data-ref="T10000"]');
  await expect(cell).toBeAttached({ timeout: 10_000 });
  const expected = SPEC137_LARGE_XLSX.t10000Value - SPEC137_LARGE_XLSX.s1Value + newS1;
  await expect
    .poll(async () => numberFromCellText((await cell.textContent()) ?? ""), { timeout: 10_000 })
    .toBeCloseTo(expected, 1);
}

async function measureSave(page: Page): Promise<{ sheetSaveMs: number; sheetMainThreadBlockDuringSaveMs: number }> {
  await resetLongTasks(page);
  const t0 = Date.now();
  const downloadPromise = page.waitForEvent("download", { timeout: 30_000 });
  await page.getByTestId("spreadsheet-download").click();
  await downloadPromise;
  const sheetSaveMs = Date.now() - t0;
  const sheetMainThreadBlockDuringSaveMs = await longestTask(page);
  return { sheetSaveMs, sheetMainThreadBlockDuringSaveMs };
}

/** Opens spec137-library.xlsx `count` times through the real open path in
 * one context, closing back to the library between opens, then hands back a
 * storageState that carries those documents' IndexedDB records into a fresh
 * context (Playwright 1.63 captures IndexedDB in storageState). */
async function seedLibrary(browser: Browser, count: number): Promise<string> {
  const context = await browser.newContext();
  const page = await context.newPage();
  await page.goto("/");
  for (let i = 0; i < count; i++) {
    await page.getByTestId("open-document-input").waitFor({ state: "attached" });
    const navigated = page.waitForURL(/\/sheets\/.+/, { timeout: 30_000 });
    await page.getByTestId("open-document-input").setInputFiles(LIBRARY_FIXTURE_PATH);
    await navigated;
    await page.getByTestId("close-document").click();
    await expect(page.getByTestId("library")).toBeVisible();
  }
  const state = await context.storageState();
  await context.close();
  return JSON.stringify(state);
}

/** Reload "/" with `storageStateJson`'s documents already in IndexedDB ->
 * performance.now() (navigation-relative) when the 20th library-row
 * attaches, plus the settled heap of that landing. */
async function measureLibraryLanding(browser: Browser, storageStateJson: string, expectedRows: number): Promise<{ sheetLibraryLandingMs: number; heapMB: number }> {
  const context = await browser.newContext({ storageState: JSON.parse(storageStateJson) });
  const page = await context.newPage();
  await page.addInitScript((n: number) => {
    const w = window as unknown as { __libraryLandingMs?: number };
    const check = () => {
      if (document.querySelectorAll('[data-testid="library-row"]').length >= n && w.__libraryLandingMs === undefined) {
        w.__libraryLandingMs = performance.now();
      }
    };
    const start = () => {
      check();
      new MutationObserver(check).observe(document.body, { childList: true, subtree: true });
    };
    if (document.body) {
      start();
    } else {
      document.addEventListener("DOMContentLoaded", start);
    }
  }, expectedRows);
  await page.goto("/");
  await page.waitForFunction(() => (window as unknown as { __libraryLandingMs?: number }).__libraryLandingMs !== undefined, { timeout: 30_000 });
  const sheetLibraryLandingMs = await page.evaluate(() => (window as unknown as { __libraryLandingMs: number }).__libraryLandingMs);
  await page.waitForTimeout(500); // let the landing settle before sampling heap
  const heapMB = await measureHeapMB(page);
  await context.close();
  return { sheetLibraryLandingMs, heapMB };
}

async function measureEmptyLibraryHeapMB(browser: Browser): Promise<number> {
  const context = await browser.newContext();
  const page = await context.newPage();
  await page.goto("/");
  await expect(page.getByTestId("library")).toBeVisible();
  await page.waitForTimeout(500);
  const heapMB = await measureHeapMB(page);
  await context.close();
  return heapMB;
}

test.describe.configure({ mode: "default" });

test.beforeAll(() => {
  buildFixturesIfMissing();
});

test("holds every SPEC-137 workbook performance budget on spec137-large.xlsx, median of 5 runs (AC-1, AC-2)", async ({ browser }) => {
  test.setTimeout(20 * 60 * 1000);
  const baseline = readBaseline();
  const samples: Record<string, number[]> = Object.fromEntries(Object.keys(BUDGETS).map((k) => [k, []]));

  // The 20 library-landing workbooks are seeded ONCE per benchmark run,
  // outside the median loop (specs/137's "The budgets": "seeded once per
  // benchmark run"); an empty library's heap is likewise stable enough
  // across runs to measure once rather than five times.
  const seeded = await seedLibrary(browser, 20);
  const emptyHeapMB = await measureEmptyLibraryHeapMB(browser);

  for (let run = 0; run < RUNS; run++) {
    const context = await browser.newContext();
    const page = await context.newPage();
    await installLongTaskCollector(page);

    const open = await measureOpen(page);
    expect(await longTasksSupported(page), "Long Tasks API available (block budgets would be vacuous without it)").toBe(true);
    samples.sheetOpenFirstPaintMs.push(open.sheetOpenFirstPaintMs);
    samples.sheetMainThreadBlockDuringOpenMs.push(open.sheetMainThreadBlockDuringOpenMs);

    const reopen = await measureReopen(page);
    samples.sheetReopenFirstPaintMs.push(reopen.sheetReopenFirstPaintMs);

    const scroll = await measureScroll(page);
    samples.sheetScrollFps.push(scroll.sheetScrollFps);
    samples.sheetMainThreadBlockDuringScrollMs.push(scroll.sheetMainThreadBlockDuringScrollMs);

    const jump = await measureJumpToEnd(page);
    samples.sheetJumpToEndMs.push(jump.sheetJumpToEndMs);

    const recalc = await measureChainRecalc(page);
    samples.sheetChainRecalcMs.push(recalc.sheetChainRecalcMs);
    samples.sheetMainThreadBlockDuringRecalcMs.push(recalc.sheetMainThreadBlockDuringRecalcMs);
    await assertChainTotal(page, recalc.newS1);

    const save = await measureSave(page);
    samples.sheetSaveMs.push(save.sheetSaveMs);
    samples.sheetMainThreadBlockDuringSaveMs.push(save.sheetMainThreadBlockDuringSaveMs);

    samples.sheetHeapAfterSettleMB.push(await measureHeapMB(page));
    await context.close();

    const landing = await measureLibraryLanding(browser, seeded, 20);
    samples.sheetLibraryLandingMs.push(landing.sheetLibraryLandingMs);
    samples.sheetLibraryHeapDeltaMB.push(landing.heapMB - emptyHeapMB);
  }

  for (const name of Object.keys(BUDGETS)) {
    assertBudget(name, samples[name], baseline);
  }
});

// AC-54: MAX_WORKBOOK_CELLS' 2,000,000 ceiling is affordable at 4x this
// fixture's quarter-of-the-ceiling size: H0 (an empty workbook's settled
// heap) + 4 * (H - H0) must stay within SPEC-047's 400MB whole-app ceiling.
test("the cell ceiling stays affordable: H0 + 4 * (H - H0) <= 400MB (AC-54)", async ({ browser }) => {
  test.setTimeout(2 * 60 * 1000);

  const emptyContext = await browser.newContext();
  const emptyPage = await emptyContext.newPage();
  await emptyPage.goto("/");
  await emptyPage.getByTestId("new-spreadsheet").first().click();
  await expect(emptyPage).toHaveURL(/\/sheets\/.+/);
  await emptyPage.waitForTimeout(500);
  const h0 = await measureHeapMB(emptyPage);
  await emptyContext.close();

  const largeContext = await browser.newContext();
  const largePage = await largeContext.newPage();
  await largePage.goto("/");
  await largePage.getByTestId("open-document-input").waitFor({ state: "attached" });
  const navigated = largePage.waitForURL(/\/sheets\/.+/, { timeout: 30_000 });
  await largePage.getByTestId("open-document-input").setInputFiles(LARGE_FIXTURE_PATH);
  await navigated;
  await waitForFirstPaintedCell(largePage);
  await largePage.waitForTimeout(2000);
  const h = await measureHeapMB(largePage);
  await largeContext.close();

  const projected = h0 + 4 * (h - h0);
  // eslint-disable-next-line no-console
  console.log(`[SPEC-137] AC-54: H0=${h0.toFixed(1)}MB H=${h.toFixed(1)}MB projected=${projected.toFixed(1)}MB (ceiling 400)`);
  expect.soft(projected, "H0 + 4 * (H - H0) <= 400MB").toBeLessThanOrEqual(400);
});
