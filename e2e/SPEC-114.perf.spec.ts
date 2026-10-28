import { test, expect, type Page } from "@playwright/test";
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-114 — search is fast enough to type into, measured in SPEC-047's
// harness: its own Playwright project (perf-library-search, the chain's tail
// in playwright.config.ts), its two gates (the spec's absolute ceiling, and
// the checked-in e2e/perf-baseline.json + 15%), medians over repeated runs.
//
// A library of LIBRARY_DOCS text documents totalling ~20 MB of text is
// written straight into IndexedDB (generated in the page from a fixed seed,
// so every run indexes the same words), indexed through the real UI — Labs
// flag on, scope Everything, "Indexed 200 of 200" — and then, over that warm
// index, every query in QUERIES is asked RUNS times through the search
// field. Each answer's round trip (asked on the main thread, searched in the
// worker, answered back; useLibrarySearch's roundTripMs, which the results
// panel carries as data-query-ms) is the sample; a query's median must be
// under the spec's 50 ms, and the gated row is the WORST query's median —
// the first keystroke's one-letter prefix included. The library must stay
// usable while it indexes: the longest main-thread task in that window is
// gated too, by the same Long Tasks API collector SPEC-047 uses.

const REPO_ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), "..");
const BASELINE_PATH = path.join(REPO_ROOT, "e2e", "perf-baseline.json");
const RUNS = 5;
const LIBRARY_DOCS = 200;
const CHARS_PER_DOC = 100_000;
const QUERIES = ["t", "th", "harbour", "harbour figures", '"the harbour"', "harbour -rotterdam", "kind:text report", "quarterly"];

const BUDGETS: Record<string, { ceiling: number; direction: "atMost" }> = {
  librarySearchQueryMs: { ceiling: 50, direction: "atMost" },
  librarySearchMainThreadBlockDuringIndexMs: { ceiling: 100, direction: "atMost" },
};

interface Baseline {
  [key: string]: number | string | undefined;
}

function readBaseline(): Baseline {
  return existsSync(BASELINE_PATH) ? (JSON.parse(readFileSync(BASELINE_PATH, "utf-8")) as Baseline) : {};
}

function median(values: number[]): number {
  const sorted = [...values].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];
}

/** Both of SPEC-047's gates on one value. */
function assertBudget(name: string, value: number, samples: number[], baseline: Baseline): void {
  const budget = BUDGETS[name];
  const baselineValue = typeof baseline[name] === "number" ? (baseline[name] as number) : undefined;
  // eslint-disable-next-line no-console
  console.log(
    `[SPEC-114] ${name}: samples=[${samples.map((n) => n.toFixed(1)).join(", ")}] value=${value.toFixed(2)} ` +
      `ceiling=<=${budget.ceiling} baseline=${baselineValue ?? "none"}`,
  );
  expect.soft(value, `${name} absolute gate (ceiling ${budget.ceiling})`).toBeLessThanOrEqual(budget.ceiling);
  if (baselineValue !== undefined) {
    expect.soft(value, `${name} relative gate (baseline ${baselineValue} + 15%)`).toBeLessThanOrEqual(baselineValue * 1.15);
  }
}

// --- Long Tasks API collector (SPEC-047's pattern, duplicated per file — no
// shared e2e helpers module exists in this repo). -----------------------------
async function installLongTaskCollector(page: Page): Promise<void> {
  await page.addInitScript(() => {
    const w = window as unknown as { __longTasks: number[] };
    w.__longTasks = [];
    try {
      new PerformanceObserver((list) => {
        for (const entry of list.getEntries()) {
          w.__longTasks.push(entry.duration);
        }
      }).observe({ type: "longtask", buffered: true });
    } catch {
      // No Long Tasks API: nothing is collected, and the row reads 0.
    }
  });
}

async function resetLongTasks(page: Page): Promise<void> {
  await page.evaluate(() => {
    (window as unknown as { __longTasks: number[] }).__longTasks = [];
  });
}

async function maxLongTask(page: Page): Promise<number> {
  return page.evaluate(() => Math.max(0, ...(window as unknown as { __longTasks: number[] }).__longTasks));
}

/** Writes the benchmark library straight into the app's IndexedDB: text
 * documents of ~charsPerDoc characters each, words drawn from a fixed
 * vocabulary by a fixed seed. */
async function seedLibrary(page: Page, docs: number, charsPerDoc: number): Promise<number> {
  return page.evaluate(
    async ({ docs, charsPerDoc }) => {
      let seed = 114;
      const rand = (): number => (seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff;
      const syllables = ["har", "bour", "fig", "ure", "quar", "ter", "ly", "re", "port", "an", "nu", "al", "mar", "ket", "rot", "dam", "ber", "lin", "plan", "ning"];
      const vocabulary = Array.from({ length: 6000 }, () => {
        let word = "";
        for (let k = 1 + Math.floor(rand() * 3); k >= 0; k--) word += syllables[Math.floor(rand() * syllables.length)];
        return word;
      });
      const common = ["the", "of", "and", "to", "in", "a", "is", "that", "for", "harbour", "figures", "report", "quarterly", "rotterdam"];
      const word = (): string => (rand() < 0.4 ? common[Math.floor(rand() * common.length)] : vocabulary[Math.floor(rand() ** 2.5 * vocabulary.length)]);
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        const request = indexedDB.open("docu-edit", 3);
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      let total = 0;
      for (let d = 0; d < docs; d++) {
        const blocks: unknown[] = [];
        let length = 0;
        while (length < charsPerDoc) {
          const text = Array.from({ length: 40 + Math.floor(rand() * 40) }, word).join(" ") + ".";
          blocks.push({ type: "paragraph", runs: [{ text, formatting: {} }] });
          length += text.length + 1;
        }
        total += length;
        const id = `perf-${String(d).padStart(3, "0")}`;
        const model = {
          id,
          title: `Benchmark ${d} ${vocabulary[d]}`,
          sections: [{ blocks }],
          defaultStyles: {
            pageSize: { widthCm: 21, heightCm: 29.7 },
            margins: { topCm: 2.54, rightCm: 2.54, bottomCm: 2.54, leftCm: 2.54 },
            fontFamily: "Calibri",
            fontSizePt: 11,
          },
        };
        await new Promise<void>((resolve, reject) => {
          const transaction = db.transaction("documents", "readwrite");
          transaction.objectStore("documents").put({ id, format: "docx", lastModified: 1_700_000_000_000 + d, model });
          transaction.oncomplete = () => resolve();
          transaction.onerror = () => reject(transaction.error);
        });
      }
      db.close();
      return total;
    },
    { docs, charsPerDoc },
  );
}

/** Types `query` into the field and returns the round trip of its answer. */
async function ask(page: Page, query: string): Promise<number> {
  const panel = page.getByTestId("library-search-panel");
  await page.getByTestId("library-search").fill(query);
  await expect(panel).toHaveAttribute("data-answered", query, { timeout: 30_000 });
  return Number(await panel.getAttribute("data-query-ms"));
}

test.describe.configure({ mode: "default" });

test("a query over a warm index of 200 documents / 20 MB of text returns in under 50 ms; indexing never blocks the library", async ({ page }) => {
  test.setTimeout(15 * 60 * 1000);
  const baseline = readBaseline();
  await installLongTaskCollector(page);
  // The library-search experiment on, as Labs stores it.
  await page.addInitScript(() => {
    localStorage.setItem("docu-edit:preferences", JSON.stringify({ experiments: { "library-search": true } }));
  });
  await page.goto("/");
  await expect(page.getByTestId("app-bar")).toBeVisible();
  const totalChars = await seedLibrary(page, LIBRARY_DOCS, CHARS_PER_DOC);
  expect(totalChars).toBeGreaterThanOrEqual(LIBRARY_DOCS * CHARS_PER_DOC);
  await page.reload();
  await expect(page.getByTestId("library-row").first()).toBeVisible({ timeout: 60_000 });

  // Build the index through the UI; the main thread's longest task while
  // it builds is the "never blocks the library" row.
  await page.waitForTimeout(1_000);
  await resetLongTasks(page);
  await page.getByTestId("search-scope").selectOption("everything");
  await expect(page.getByTestId("search-index-progress")).toHaveText(`Indexed ${LIBRARY_DOCS} of ${LIBRARY_DOCS}`, { timeout: 10 * 60 * 1000 });
  const blockDuringIndex = await maxLongTask(page);

  // Warm: every query once, unmeasured, then RUNS measured rounds.
  for (const query of QUERIES) {
    await ask(page, query);
  }
  const samples: Record<string, number[]> = Object.fromEntries(QUERIES.map((query) => [query, []]));
  for (let run = 0; run < RUNS; run++) {
    for (const query of QUERIES) {
      samples[query].push(await ask(page, query));
    }
  }
  const medians = QUERIES.map((query) => median(samples[query]));
  QUERIES.forEach((query, i) => {
    // eslint-disable-next-line no-console
    console.log(`[SPEC-114] ${JSON.stringify(query)}: samples=[${samples[query].map((n) => n.toFixed(1)).join(", ")}] median=${medians[i].toFixed(2)}`);
  });
  assertBudget("librarySearchQueryMs", Math.max(...medians), medians, baseline);
  assertBudget("librarySearchMainThreadBlockDuringIndexMs", blockDuringIndex, [blockDuringIndex], baseline);
  // The answers are real: the common words are in every document.
  await page.getByTestId("library-search").fill("harbour");
  await expect(page.getByTestId("search-result")).toHaveCount(50);
  await expect(page.getByTestId("search-show-more")).toBeVisible();
});
