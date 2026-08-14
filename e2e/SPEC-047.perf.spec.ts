import { test, expect, type Page } from "@playwright/test";
import { existsSync, readFileSync } from "node:fs";
import { execFileSync } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { MOD } from "./keys";

// SPEC-047 — large document performance. The deterministic benchmark: every
// row of specs/047-large-document-performance.md's "The budgets" table,
// asserted against the generated fixture (scripts/build-spec047-fixture.mjs)
// with the two-gate rule from "The regression gate" (absolute ceiling,
// relative to the checked-in e2e/perf-baseline.json). No shared e2e helpers
// module exists in this repo (see the other spec files) — every helper below
// is local to this file, matching that convention.
//
// Timing validity: the baselines are calibrated for THIS FILE running with
// the machine otherwise quiet. A full-suite `npx playwright test` runs
// sibling spec files in parallel workers, whose CPU load can inflate these
// samples — the median-of-5 plus the 15% relative allowance absorb typical
// contention, but a timing failure seen only in a full-suite run should be
// confirmed with a solo run of this file before being treated as a
// regression. The same applies at HOST level: the factory's gate container
// shares its host machine's CPU, so heavy work on the host during a gate
// run (builds, other browser suites, review agents) starves the container
// and inflates every row together — a gate failure whose rows all drifted
// ~2x in unison is contention, and a quiet re-run is the diagnostic.

const REPO_ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), "..");
const FIXTURE_PATH = path.join(REPO_ROOT, "e2e", "fixtures", "spec-047-fixture.docx");
const BASELINE_PATH = path.join(REPO_ROOT, "e2e", "perf-baseline.json");
const RUNS = 5;

// --- "The budgets" table, restated as data so AC-1 ("each row is asserted
// by the benchmark") is checkable by reading this object against the
// table, not by re-deriving it from scattered assertions. ------------------
const BUDGETS: Record<string, { ceiling: number; direction: "atMost" | "atLeast" }> = {
  firstPageMs: { ceiling: 2000, direction: "atMost" },
  fullyInteractiveMs: { ceiling: 8000, direction: "atMost" },
  typingLatencyP95Ms: { ceiling: 50, direction: "atMost" },
  scrollFps: { ceiling: 30, direction: "atLeast" },
  mainThreadBlockDuringSaveMs: { ceiling: 100, direction: "atMost" },
  mainThreadBlockDuringScrollMs: { ceiling: 100, direction: "atMost" },
  mainThreadBlockDuringTypingMs: { ceiling: 100, direction: "atMost" },
  mainThreadBlockDuringOpenMs: { ceiling: 150, direction: "atMost" },
  saveMs: { ceiling: 5000, direction: "atMost" },
  heapAfterSettleMB: { ceiling: 400, direction: "atMost" },
};

// The 30-minute leak test (AC-18) is its own single-run check, not part of
// the median-of-5 loop above: running a 30-minute session five times would
// take over two hours, which is not what "a benchmark run by `npx
// playwright test`" can mean. Its ceiling is the 10% growth allowance
// itself, not a baseline-relative gate.
const HEAP_GROWTH_CEILING_PERCENT = 10;

function buildFixtureIfMissing(): void {
  if (existsSync(FIXTURE_PATH)) {
    return;
  }
  execFileSync("npx", ["vite-node", "scripts/build-spec047-fixture.mjs"], {
    cwd: REPO_ROOT,
    stdio: "inherit",
  });
  if (!existsSync(FIXTURE_PATH)) {
    throw new Error(`Fixture build did not produce ${FIXTURE_PATH}`);
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

/** AC-22 (absolute) + AC-23 (relative, 15% over the checked-in baseline) —
 * both evaluated on the median of `samples`. `expect.soft` so one budget's
 * failure doesn't hide another's in the same run (AC-1: every row is
 * actually asserted, every run). */
function assertBudget(name: string, samples: number[], baseline: Baseline): void {
  const budget = BUDGETS[name];
  const value = median(samples);
  const baselineValue = typeof baseline[name] === "number" ? (baseline[name] as number) : undefined;
  // eslint-disable-next-line no-console
  console.log(
    `[SPEC-047] ${name}: samples=[${samples.map((n) => n.toFixed(1)).join(", ")}] median=${value.toFixed(2)} ` +
      `ceiling=${budget.direction === "atMost" ? "<=" : ">="}${budget.ceiling} baseline=${baselineValue ?? "none"}`,
  );
  if (budget.direction === "atMost") {
    expect.soft(value, `${name} absolute gate (ceiling ${budget.ceiling})`).toBeLessThanOrEqual(budget.ceiling);
    if (baselineValue !== undefined) {
      expect
        .soft(value, `${name} relative gate (baseline ${baselineValue} + 15%)`)
        .toBeLessThanOrEqual(baselineValue * 1.15);
    }
  } else {
    expect.soft(value, `${name} absolute gate (floor ${budget.ceiling})`).toBeGreaterThanOrEqual(budget.ceiling);
    if (baselineValue !== undefined) {
      expect
        .soft(value, `${name} relative gate (baseline ${baselineValue} - 15%)`)
        .toBeGreaterThanOrEqual(baselineValue * 0.85);
    }
  }
}

// --- Long Tasks API collector: installed once per page (persists across
// the client-side route change from "/" to "/doc/:id" — no full reload in
// between), reset at the start of each phase we want an isolated
// "longest task" sample for. -------------------------------------------
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
      // Support is asserted by the budget test: an "atMost" block budget
      // reads 0 when the observer never fires, and a 0 from "no long tasks"
      // must be distinguishable from a 0 from "the API never worked" — the
      // latter would pass every main-thread-block gate vacuously.
      w.__longTasksSupported = true;
    } catch {
      // Left false; the budget test fails loudly on it.
    }
  });
}

async function longTasksSupported(page: Page): Promise<boolean> {
  return page.evaluate(
    () => (window as unknown as { __longTasksSupported?: boolean }).__longTasksSupported === true,
  );
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

/** CDP `Performance.getMetrics` -> `JSHeapUsedSize`, sampled after
 * `HeapProfiler.collectGarbage` (AC-17): an unforced sample measures
 * garbage that has not been collected yet, not memory actually held. */
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

async function waitForFullyInteractive(page: Page, timeout = 30_000): Promise<void> {
  await expect(page.getByTestId("document-opening")).toBeHidden({ timeout });
  await expect(page.getByTestId("document-pages")).toHaveAttribute("data-pagination-state", "clean", { timeout });
}

/** Opens the fixture through the real "Open document" picker path (not an
 * IndexedDB seed) and waits until it is fully interactive — the shared tail
 * every measurement below builds on. */
async function openFixture(page: Page): Promise<void> {
  await page.goto("/");
  const navigated = page.waitForURL(/\/doc\/.+/, { timeout: 30_000 });
  await page.getByTestId("open-document-input").setInputFiles(FIXTURE_PATH);
  await navigated;
  await waitForFullyInteractive(page);
}

interface OpenSample {
  firstPageMs: number;
  fullyInteractiveMs: number;
  mainThreadBlockDuringOpenMs: number;
  openingIndicatorWasVisible: boolean;
}

/** AC-3, AC-4, AC-5, AC-16: opens the fixture, timing from navigation (the
 * AC-3 definition: "Date.now() from navigation to first
 * [data-testid=document-page] with content") through first-page-visible and
 * on to fully interactive, while also recording whether the open-progress
 * indicator was ever actually shown along the way. */
async function measureOpen(page: Page): Promise<OpenSample> {
  await page.goto("/");
  await resetLongTasks(page);
  const navigated = page.waitForURL(/\/doc\/.+/, { timeout: 30_000 });
  // t0 at the moment the file is chosen, NOT at the URL change: in this SPA
  // every expensive part of opening (read, unzip, parse, store) happens
  // BEFORE router.push, so "from navigation" — the spec's original wording,
  // amended at gate — excluded exactly the phase AC-3's 2s budget exists to
  // bound.
  const t0 = Date.now();
  await page.getByTestId("open-document-input").setInputFiles(FIXTURE_PATH);
  await navigated;

  await page.waitForFunction(
    () => {
      const pages = document.querySelectorAll('[data-testid="document-page"]');
      for (const el of pages) {
        if ((el.textContent ?? "").trim().length > 0) {
          return true;
        }
      }
      return false;
    },
    { timeout: 15_000 },
  );
  const firstPageMs = Date.now() - t0;

  const openingIndicatorWasVisible = await page
    .getByTestId("document-opening")
    .isVisible()
    .catch(() => false);

  await waitForFullyInteractive(page);
  const fullyInteractiveMs = Date.now() - t0;

  const mainThreadBlockDuringOpenMs = await longestTask(page);
  return { firstPageMs, fullyInteractiveMs, mainThreadBlockDuringOpenMs, openingIndicatorWasVisible };
}

/** AC-8, AC-15: a 5-second programmatic sweep down and back up the
 * document, sampling animation-frame timestamps for fps and the Long Tasks
 * API for the phase's longest main-thread block. */
async function measureScroll(page: Page): Promise<{ scrollFps: number; mainThreadBlockDuringScrollMs: number }> {
  await resetLongTasks(page);
  const scrollFps = await page.evaluate(() => {
    const container = document.querySelector("main");
    if (!container) {
      return 0;
    }
    const durationMs = 5000;
    // A bounded, constant-rate sweep (down for the first half, back up for
    // the second — exercises page recycling in both directions, AC-9)
    // rather than the full document height: sweeping all ~277 pages in 5s
    // would mean ~55 pages/s, far faster than any real scroll gesture and
    // not what "does scrolling feel smooth" is asking about. 4000px/s is a
    // fast, still-plausible mouse-wheel/trackpad rate.
    const speedPxPerSec = 4000;
    const maxScroll = Math.max(0, container.scrollHeight - container.clientHeight);
    const sweepDistance = Math.min(maxScroll, speedPxPerSec * (durationMs / 1000 / 2));
    return new Promise<number>((resolve) => {
      const start = performance.now();
      let frames = 0;
      function step() {
        frames++;
        const elapsed = performance.now() - start;
        const t = Math.min(1, elapsed / durationMs);
        const pos = t < 0.5 ? t * 2 * sweepDistance : (1 - (t - 0.5) * 2) * sweepDistance;
        container!.scrollTop = pos;
        if (elapsed < durationMs) {
          requestAnimationFrame(step);
        } else {
          resolve(frames / (elapsed / 1000));
        }
      }
      requestAnimationFrame(step);
    });
  });
  const mainThreadBlockDuringScrollMs = await longestTask(page);
  return { scrollFps, mainThreadBlockDuringScrollMs };
}

/** AC-6, AC-15: 200 keystrokes at the document end, keydown to
 * model-updated paint — a `keydown` capture listener stamps the start,
 * a `MutationObserver` reacting to the resulting DOM change stamps the end
 * one animation frame later (the paint proxy Playwright's own event loop
 * can observe). */
async function measureTyping(
  page: Page,
): Promise<{ typingLatencyP95Ms: number; mainThreadBlockDuringTypingMs: number }> {
  await page.getByTestId("document-content").click();
  await page.keyboard.press(`${MOD}+End`);
  await resetLongTasks(page);

  await page.evaluate(() => {
    const w = window as unknown as { __typingLatencies: number[]; __typingObserver?: MutationObserver };
    w.__typingLatencies = [];
    const root = document.querySelector('[data-testid="document-content"]');
    if (!root) {
      return;
    }
    let pendingStart: number | null = null;
    root.addEventListener(
      "keydown",
      () => {
        pendingStart = performance.now();
      },
      true,
    );
    const observer = new MutationObserver(() => {
      if (pendingStart !== null) {
        const start = pendingStart;
        pendingStart = null;
        requestAnimationFrame(() => {
          w.__typingLatencies.push(performance.now() - start);
        });
      }
    });
    observer.observe(root, { childList: true, subtree: true, characterData: true });
    w.__typingObserver = observer;
  });

  for (let i = 0; i < 200; i++) {
    await page.keyboard.press("x");
  }
  // Lets the final keystroke's rAF-scheduled sample land before reading back.
  await page.waitForTimeout(100);

  const latencies = await page.evaluate(() => {
    const w = window as unknown as { __typingLatencies: number[]; __typingObserver?: MutationObserver };
    w.__typingObserver?.disconnect();
    return w.__typingLatencies;
  });
  // Guard against a vacuously green p95: if typing silently stopped working
  // (a regression that leaves the document read-only, say), zero samples
  // would sort to a p95 of 0 and "pass". Most of the 200 keystrokes must
  // actually have been observed mutating the DOM.
  expect(latencies.length, "typing latency samples actually collected").toBeGreaterThan(100);
  const sorted = [...latencies].sort((a, b) => a - b);
  const p95Index = Math.min(sorted.length - 1, Math.max(0, Math.ceil(sorted.length * 0.95) - 1));
  const typingLatencyP95Ms = sorted.length > 0 ? sorted[p95Index] : 0;

  const mainThreadBlockDuringTypingMs = await longestTask(page);
  return { typingLatencyP95Ms, mainThreadBlockDuringTypingMs };
}

/** AC-19, AC-20: invoke-to-Blob-resolved timing (the 'download' event fires
 * once the object URL is ready) plus the phase's longest main-thread block
 * (AC-20's "does not freeze the UI", made concrete). */
async function measureSave(page: Page): Promise<{ saveMs: number; mainThreadBlockDuringSaveMs: number }> {
  await resetLongTasks(page);
  const t0 = Date.now();
  await page.getByTestId("download-menu").click();
  const downloadPromise = page.waitForEvent("download", { timeout: 30_000 });
  await page.getByTestId("download-docx").click();
  await downloadPromise;
  const saveMs = Date.now() - t0;
  const mainThreadBlockDuringSaveMs = await longestTask(page);
  return { saveMs, mainThreadBlockDuringSaveMs };
}

// Registration order, not parallel across workers (this file's tests all
// exercise the same fixture and compete for the same CPU a timing budget
// cares about) — but NOT "serial" mode: that aborts every later test the
// moment one fails, and these tests are independent (AC-7's DOM-count check
// has no reason to go unreported just because a save-latency budget did).
test.describe.configure({ mode: "default" });

test.beforeAll(() => {
  buildFixtureIfMissing();
});

test("holds every SPEC-047 performance budget on the fixture, median of 5 runs (AC-1..6, AC-15, AC-16, AC-17, AC-19, AC-20, AC-22, AC-23)", async ({
  browser,
}) => {
  test.setTimeout(15 * 60 * 1000);
  const baseline = readBaseline();

  const samples: Record<string, number[]> = Object.fromEntries(Object.keys(BUDGETS).map((k) => [k, []]));
  let openingIndicatorEverVisible = false;

  for (let run = 0; run < RUNS; run++) {
    const context = await browser.newContext();
    const page = await context.newPage();
    await installLongTaskCollector(page);

    const open = await measureOpen(page);
    // Every main-thread-block budget below reads the Long Tasks collector;
    // a 0 is only meaningful when the observer actually installed.
    expect(await longTasksSupported(page), "Long Tasks API available (block budgets would be vacuous without it)").toBe(
      true,
    );
    samples.firstPageMs.push(open.firstPageMs);
    samples.fullyInteractiveMs.push(open.fullyInteractiveMs);
    samples.mainThreadBlockDuringOpenMs.push(open.mainThreadBlockDuringOpenMs);
    openingIndicatorEverVisible ||= open.openingIndicatorWasVisible;

    const scroll = await measureScroll(page);
    samples.scrollFps.push(scroll.scrollFps);
    samples.mainThreadBlockDuringScrollMs.push(scroll.mainThreadBlockDuringScrollMs);

    const typing = await measureTyping(page);
    samples.typingLatencyP95Ms.push(typing.typingLatencyP95Ms);
    samples.mainThreadBlockDuringTypingMs.push(typing.mainThreadBlockDuringTypingMs);

    const save = await measureSave(page);
    samples.saveMs.push(save.saveMs);
    samples.mainThreadBlockDuringSaveMs.push(save.mainThreadBlockDuringSaveMs);

    samples.heapAfterSettleMB.push(await measureHeapMB(page));

    await context.close();
  }

  for (const name of Object.keys(BUDGETS)) {
    assertBudget(name, samples[name], baseline);
  }

  // AC-5: the open-progress indicator must actually have been visible at
  // some point while opening, not just architecturally present.
  expect.soft(openingIndicatorEverVisible, "AC-5: open-progress indicator was visible during at least one run").toBe(
    true,
  );
});

test("only pages inside the realization window are rendered to the DOM (AC-2, AC-7)", async ({ page }) => {
  await installLongTaskCollector(page);
  await openFixture(page);

  const totalPages = await page.locator('[data-testid="document-page"]').count();
  expect(totalPages, "the fixture is large enough to need virtualization").toBeGreaterThan(100);

  const domBlocks = await page.locator('[data-testid="document-page"] [data-section][data-block]').count();
  // Same approach as SPEC-028.spec.ts:687's <1000 ceiling for its 3000-block
  // document, re-sized for this fixture: its realization window carries
  // denser pages (tables contribute a row element per data-block address),
  // so the equivalent "only the window is realized" bound lands at 1500.
  expect(domBlocks, "only realized pages' blocks are in the DOM").toBeLessThan(1500);
});

test("scroll position is stable when off-screen pages are recycled (AC-9)", async ({ page }) => {
  await installLongTaskCollector(page);
  await openFixture(page);

  const target = await page.evaluate(() => {
    const container = document.querySelector("main")!;
    container.scrollTop = 5000;
    return container.scrollTop;
  });
  // Scroll far away (recycling the pages around `target`) and back.
  await page.evaluate(() => {
    document.querySelector("main")!.scrollTop = 40_000;
  });
  await page.waitForTimeout(300);
  await page.evaluate((pos) => {
    document.querySelector("main")!.scrollTop = pos;
  }, target);
  await page.waitForTimeout(300);
  const restored = await page.evaluate(() => document.querySelector("main")!.scrollTop);
  expect(restored).toBe(target);
});

// AC-18: a 30-minute scripted editing session (>200 edits), sampled once a
// minute, ending within 110% of the first post-settle heap sample. A
// single run, not the median-of-5 loop above — five 30-minute sessions
// would take over two hours, which is not what running this benchmark can
// mean in practice. Skipped by default (30+ minutes is far past a routine
// pass) but deliberately runnable without editing source:
// `SPEC047_LEAK=1 npx playwright test e2e/SPEC-047.perf.spec.ts -g "30-minute"`.
test("a 30-minute editing session keeps heap growth within 10% of the first post-settle sample (AC-18)", async ({
  page,
}) => {
  test.skip(!process.env.SPEC047_LEAK, "set SPEC047_LEAK=1 to run the 30-minute leak session");
  test.setTimeout(34 * 60 * 1000);
  await installLongTaskCollector(page);
  await openFixture(page);

  await page.getByTestId("document-content").click();
  await page.keyboard.press(`${MOD}+End`);

  // Prime the undo history past its MAX_HISTORY_STEPS = 200 cap BEFORE the
  // baseline sample: the spec's own rationale for the >200-edit floor is
  // that "below that cap, growth is the feature working. Above it, the
  // stack evicts and the curve must go flat" — so the flat-curve assertion
  // only means something once the legitimate fill has already happened.
  // Measured against this fixture, the fill adds ~50% to the post-settle
  // heap (each step retains a cloned blocks array); a baseline taken before
  // it mistakes the feature for the leak, which is exactly what the spec
  // says not to do. Type/Backspace pairs alternate step labels, so they
  // never coalesce: 130 pairs = 260 steps > 200.
  for (let i = 0; i < 130; i++) {
    await page.keyboard.type("x");
    await page.keyboard.press("Backspace");
  }
  await page.waitForTimeout(1000);

  const firstHeapMB = await measureHeapMB(page);
  let lastHeapMB = firstHeapMB;

  const SESSION_MS = 30 * 60 * 1000;
  const SAMPLE_INTERVAL_MS = 60 * 1000;
  const EDIT_INTERVAL_MS = 1000; // >200 edits across 30 minutes (spec's own "more than 200" floor).

  const start = Date.now();
  let lastSample = start;
  let editCount = 0;
  while (Date.now() - start < SESSION_MS) {
    await page.keyboard.type("x");
    await page.keyboard.press("Backspace");
    editCount += 2;
    if (Date.now() - lastSample >= SAMPLE_INTERVAL_MS) {
      lastHeapMB = await measureHeapMB(page);
      lastSample = Date.now();
    }
    await page.waitForTimeout(EDIT_INTERVAL_MS);
  }
  lastHeapMB = await measureHeapMB(page);

  expect(editCount, "more than 200 edits during the measured window itself, so eviction stays active").toBeGreaterThan(
    200,
  );
  expect
    .soft(lastHeapMB, `AC-18: last heap sample within ${HEAP_GROWTH_CEILING_PERCENT}% of the post-cap baseline sample`)
    .toBeLessThanOrEqual(firstHeapMB * (1 + HEAP_GROWTH_CEILING_PERCENT / 100));
});
