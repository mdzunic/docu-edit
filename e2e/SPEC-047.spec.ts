import { test, expect, type Page } from "@playwright/test";
import { MOD } from "./keys";

// SPEC-047 — large document performance: the supported-ceiling read-only
// mode (AC-26, AC-27). The performance budgets themselves are asserted in
// e2e/SPEC-047.perf.spec.ts against the generated fixture; this file covers
// the functional behaviour at the ceiling, which does not need the full
// fixture to exercise.

/** Writes a committed document straight into the app's store with `count`
 * plain paragraph blocks — same shape/precedent as SPEC-010's
 * injectLargeDocument, sized here to cross the 25,000-block ceiling
 * (SPEC-047's "Blocks (paragraphs, tables, images)" limit) without needing a
 * real .docx parse. */
async function injectBlockCountDocument(
  page: Page,
  id: string,
  count: number,
  options: { longParagraphs?: boolean } = {},
): Promise<void> {
  await page.goto("/");
  await page.evaluate(
    async ([docId, blockCount, longParagraphs]) => {
      // longParagraphs: varied multi-line text per block, so canvas
      // measurement has real wrap work per paragraph and the measure pass
      // takes long enough to interact during (the focus-guard test below
      // needs that window to be reliably wide).
      const filler =
        "The quick brown fox jumps over the lazy dog while the measurement pass earns its keep. ";
      const blocks = Array.from({ length: blockCount as number }, (_, i) => ({
        type: "paragraph",
        runs: [
          {
            text: longParagraphs
              ? `Paragraph ${i} ${filler.repeat(1 + (i % 4))}`
              : `Paragraph ${i}`,
            formatting: {},
          },
        ],
      }));
      const record = {
        id: docId,
        format: "new",
        lastModified: Date.now(),
        model: {
          id: docId,
          title: "Ceiling test document",
          sections: [{ blocks }],
          defaultStyles: {
            pageSize: { widthCm: 21, heightCm: 29.7 },
            margins: { topCm: 2.54, rightCm: 2.54, bottomCm: 2.54, leftCm: 2.54 },
            fontFamily: "Calibri",
            fontSizePt: 11,
          },
        },
      };
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        const request = indexedDB.open("docu-edit");
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      await new Promise<void>((resolve, reject) => {
        const tx = db.transaction("documents", "readwrite");
        tx.objectStore("documents").put(record);
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
      });
      db.close();
    },
    [id, count, options.longParagraphs === true] as const,
  );
}

test("a document past the block ceiling opens read-only with a warning naming the limit and value (AC-26, AC-27)", async ({
  page,
}) => {
  await injectBlockCountDocument(page, "over-ceiling-doc", 25_001);
  await page.goto("/doc/over-ceiling-doc");

  const bar = page.getByTestId("too-large-bar");
  await expect(bar).toBeVisible();
  await expect(bar).toHaveText(
    "This document is too large to edit (25,001 blocks; the limit is 25,000). It's open for reading only.",
  );

  // Editing commands are unavailable: the canvas is not editable, and typing
  // does nothing.
  await expect(page.getByTestId("document-content")).toHaveAttribute("contenteditable", "false");
  await page.getByTestId("document-content").click();
  await page.keyboard.type("should not appear");
  await expect(page.getByTestId("document-content")).not.toContainText("should not appear");

  // Opening, scrolling, find, print, PDF export and download still work.
  await expect(page.getByTestId("document-content")).toContainText("Paragraph 0");
  await page.mouse.wheel(0, 2000);

  await page.keyboard.press(`${MOD}+f`);
  await expect(page.getByTestId("find-query")).toBeVisible();
  await page.getByTestId("find-query").fill("Paragraph 5");
  await expect(page.getByTestId("find-count")).not.toHaveText("0 of 0");
  await page.keyboard.press("Escape");

  await page.getByTestId("download-menu").click();
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("download-docx").click();
  const download = await downloadPromise;
  expect(download.suggestedFilename()).toMatch(/\.docx$/);
});

test("a document within every limit opens editable with no warning", async ({ page }) => {
  await injectBlockCountDocument(page, "within-ceiling-doc", 5);
  await page.goto("/doc/within-ceiling-doc");

  await expect(page.getByTestId("too-large-bar")).toBeHidden();
  await expect(page.getByTestId("document-content")).toHaveAttribute("contenteditable", "true");
  await page.getByTestId("document-content").click();
  await page.keyboard.type("editable");
  await expect(page.getByTestId("document-content")).toContainText("editable");
});

test("the deferred session restore never steals focus from a dialog opened while the first layout is pending (AC-16 regression)", async ({
  page,
}) => {
  // The session restore waits for the deferred first layout (see
  // pagination.prime()) before focusing the canvas, and must not take
  // focus if the user moved it somewhere else while it waited. The
  // steal window scales with storage/lock latency, so this scenario is
  // deterministic only under load — SPEC-021's suite inside a full run is
  // where a regression reliably surfaces; this test pins the interaction
  // itself (dialog opened during open; focus intact once settled) and
  // never false-fails on a fast machine.
  await injectBlockCountDocument(page, "focus-guard-doc", 12_000, { longParagraphs: true });
  await page.goto("/doc/focus-guard-doc");

  // Open the shortcuts dialog at the earliest moment the registry is live —
  // pressing in a retry loop lands within the deferred-layout window
  // without depending on catching the opening indicator mid-flash. If a
  // fast run settles before the dialog opens, every assertion below still
  // holds — the test just exercises the guard less deeply that run.
  await expect(async () => {
    await page.keyboard.press(`${MOD}+/`);
    await expect(page.getByTestId("shortcuts-dialog")).toBeVisible({ timeout: 150 });
  }).toPass({ timeout: 15_000 });
  await expect(page.getByTestId("shortcuts-search")).toBeFocused();

  // Once pagination settles, the restore has run (or declined to): focus
  // must still be in the dialog, not the canvas.
  await expect(page.getByTestId("document-pages")).toHaveAttribute("data-pagination-state", "clean", {
    timeout: 30_000,
  });
  await expect(page.getByTestId("shortcuts-search")).toBeFocused();
});
