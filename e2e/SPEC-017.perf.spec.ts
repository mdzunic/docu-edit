import { test, expect, type Page } from "@playwright/test";
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { MOD } from "./keys";

// SPEC-017 — find and replace: AC-30's large-document performance contract.
//
// This test used to live in e2e/SPEC-017.spec.ts and ran inside the main
// `chromium` project, fullyParallel with every functional test. Its baseline
// (perf-baseline.json's findFirstMatchMs) is calibrated for a quiet machine,
// so being dispatched mid-suite timed the sibling workers' contention as much
// as the search itself. As the suite grew that contention crept the median
// over the 15% relative gate on a clean container clone. Moving it into its
// own late-running project (playwright.config.ts, alongside SPEC-047/SPEC-052)
// runs it once the functional tests are dispatched — the quiet window its
// baseline assumes — without changing a single assertion. See the SPEC-047
// perf file header for the shared calibration rationale.

async function createDocument(page: Page): Promise<string> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  const url = page.url();
  return url.slice(url.lastIndexOf("/") + 1);
}

const content = (page: Page) => page.getByTestId("document-content");
const findQuery = (page: Page) => page.getByTestId("find-query");

// --- AC-30's "first results within 300 ms" budget ------------------------
// The 300ms is the spec's own number, so it stays asserted — but sampled the
// way e2e/SPEC-047.perf.spec.ts samples its budgets, because a single sample
// in a fullyParallel suite (playwright.config.ts, retries: 0) times this
// machine's spare CPU as much as it times the search. The median of one
// sample per needle below is gated twice: against the spec's absolute
// ceiling, and against the container-calibrated findFirstMatchMs row of
// e2e/perf-baseline.json.
//
// One needle per sample, all distinct: the search recomputes only when the
// query changes, and each of these words appears exactly once in every
// seeded paragraph, so all five samples scan the same document for the same
// number of matches.
//
// Timing validity, as in SPEC-047: the baseline row is calibrated for the
// factory container with the machine otherwise quiet. A timing failure seen
// only in a full-suite run should be confirmed by re-running this file solo
// before it is treated as a regression.
const FIND_PERF_NEEDLES = ["needle", "target", "scattered", "words", "about"];
const FIND_FIRST_MATCH_CEILING_MS = 300;
const BASELINE_PATH = path.join(
  path.dirname(fileURLToPath(import.meta.url)),
  "perf-baseline.json",
);

function median(values: number[]): number {
  const sorted = [...values].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];
}

/** The checked-in findFirstMatchMs baseline, or undefined when the file is
 * absent — the absolute ceiling above is what makes AC-30 pass or fail, and
 * it never depends on the baseline being present. */
function readBaselineFindMs(): number | undefined {
  if (!existsSync(BASELINE_PATH)) {
    return undefined;
  }
  const baseline = JSON.parse(readFileSync(BASELINE_PATH, "utf-8")) as Record<string, unknown>;
  return typeof baseline.findFirstMatchMs === "number" ? baseline.findFirstMatchMs : undefined;
}

test("a 200-page document returns first results within 300ms and typing stays responsive (AC-30, AC-31)", async ({
  page,
}) => {
  test.setTimeout(120_000);
  const docId = await createDocument(page);
  await content(page).click();
  await page.keyboard.type("seed");
  await expect(page.getByTestId("save-state")).toHaveText("All changes saved");

  // Navigate off the document first so its autosave (beforeunload/pagehide)
  // can't race the direct IndexedDB seed below and reintroduce the small
  // in-memory model; the document-list route shares the origin/IndexedDB
  // but has no autosave tied to this document.
  await page.goto("/");
  const seeded = await page.evaluate(async (id) => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const req = indexedDB.open("docu-edit"); // no explicit version: the app has already opened the DB at its current schema version by this point
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
    const readTx = db.transaction("documents", "readonly");
    const getRq = readTx.objectStore("documents").get(id);
    const committed: { model: { sections: { blocks: unknown[] }[] } } = await new Promise((resolve, reject) => {
      getRq.onsuccess = () => resolve(getRq.result);
      getRq.onerror = () => reject(getRq.error);
    });
    const paragraphs = [];
    // ~45 lines/page * 200 pages of body text.
    for (let i = 0; i < 9000; i++) {
      paragraphs.push({
        runs: [{ text: `Paragraph number ${i} with some target needle words scattered about.`, formatting: {} }],
      });
    }
    committed.model.sections[0].blocks = paragraphs;
    const writeTx = db.transaction("documents", "readwrite");
    const store = writeTx.objectStore("documents");
    store.put({ ...committed, lastModified: Date.now() });
    store.delete("draft:" + id);
    await new Promise((resolve, reject) => {
      writeTx.oncomplete = () => resolve(undefined);
      writeTx.onerror = () => reject(writeTx.error);
    });
    const verifyRq = db.transaction("documents", "readonly").objectStore("documents").get(id);
    const verified: { model: { sections: { blocks: unknown[] }[] } } = await new Promise((resolve, reject) => {
      verifyRq.onsuccess = () => resolve(verifyRq.result);
      verifyRq.onerror = () => reject(verifyRq.error);
    });
    db.close();
    return verified.model.sections[0].blocks.length;
  }, docId);
  expect(seeded).toBe(9000);

  await page.goto(`/doc/${docId}`);
  await expect(content(page).locator(".doc-paragraph").first()).toBeVisible();
  await content(page).click();

  await page.keyboard.press(`${MOD}+f`);
  const samples: number[] = [];
  for (const needle of FIND_PERF_NEEDLES) {
    await findQuery(page).fill("");
    await page.waitForFunction(() => document.querySelector('[data-testid="find-count"]')?.textContent === "0 of 0");
    // Drain the previous sample's idle-scheduled highlight measuring
    // (useFindReplace's scheduleFindMeasure) so it cannot land inside the
    // next sample's window.
    await page.evaluate(
      () => new Promise<void>((resolve) => {
        requestIdleCallback(() => resolve(), { timeout: 500 });
      }),
    );
    const start = Date.now();
    await findQuery(page).fill(needle);
    // waitForFunction polls per animation frame; expect().not.toHaveText()
    // polls on a ~100ms ladder, which quantizes a 300ms budget badly enough
    // to matter on its own.
    await page.waitForFunction(
      () => document.querySelector('[data-testid="find-count"]')?.textContent !== "0 of 0",
      { timeout: 10_000 },
    );
    samples.push(Date.now() - start);
  }
  const firstMatchMs = median(samples);
  // eslint-disable-next-line no-console
  console.log(
    `[SPEC-017] findFirstMatchMs: samples=[${samples.join(", ")}] median=${firstMatchMs} ` +
      `ceiling=<=${FIND_FIRST_MATCH_CEILING_MS} baseline=${readBaselineFindMs() ?? "none"}`,
  );
  expect
    .soft(firstMatchMs, `find first-match absolute gate (AC-30 ceiling ${FIND_FIRST_MATCH_CEILING_MS}ms)`)
    .toBeLessThanOrEqual(FIND_FIRST_MATCH_CEILING_MS);
  const baselineMs = readBaselineFindMs();
  if (baselineMs !== undefined) {
    expect
      .soft(firstMatchMs, `find first-match relative gate (baseline ${baselineMs}ms + 15%)`)
      .toBeLessThanOrEqual(baselineMs * 1.15);
  }

  // Typing stays responsive: a burst of keystrokes into the query field is
  // never dropped while the (much more expensive) search/highlight work for
  // a 9000-paragraph document runs behind it.
  await findQuery(page).fill("");
  const typed = "abcdefghijklmnopqrstuvwxyz";
  await findQuery(page).pressSequentially(typed, { delay: 10 });
  await expect(findQuery(page)).toHaveValue(typed);
});
