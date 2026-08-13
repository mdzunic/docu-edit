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
async function injectBlockCountDocument(page: Page, id: string, count: number): Promise<void> {
  await page.goto("/");
  await page.evaluate(
    async ([docId, blockCount]) => {
      const blocks = Array.from({ length: blockCount as number }, (_, i) => ({
        type: "paragraph",
        runs: [{ text: `Paragraph ${i}`, formatting: {} }],
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
    [id, count] as const,
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
