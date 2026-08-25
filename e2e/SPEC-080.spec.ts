import { test, expect, type Page } from "@playwright/test";
import { tmpdir } from "node:os";
import { join } from "node:path";

// SPEC-080 — create a blank presentation from the landing page and manage its
// slides: add, duplicate, reorder by keyboard, delete, undo the delete, then
// download and re-open the file and confirm the filmstrip's slide count.
// Placeholders cannot be typed into yet (out of scope), so slides carry no
// visible title — the deck's identity here is its slide count and numbered
// thumbnails, which is what the round trip must preserve.

async function newPresentationFromLanding(page: Page): Promise<void> {
  await page.goto("/");
  // The first-run empty state offers no "New presentation"; create a document
  // to populate the library, then the library controls expose it (AC-1).
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  await page.goto("/");
  await page.getByTestId("new-presentation").click();
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();
}

test("create a deck, manage slides, undo a delete, and round-trip the download", async ({ page }) => {
  await newPresentationFromLanding(page);
  const thumbs = page.getByTestId("filmstrip-slide");
  await expect(thumbs).toHaveCount(1);

  // Add a slide (becomes current) → 2.
  await page.getByTestId("presentation-add-slide").click();
  await expect(thumbs).toHaveCount(2);

  // Duplicate the current (2nd) slide via its ⋯ menu → 3.
  await page.getByTestId("presentation-slide-menu").nth(1).click();
  await page.getByTestId("slide-menu-duplicate").click();
  await expect(thumbs).toHaveCount(3);

  // Reorder by keyboard: open the current thumbnail's menu with the Menu key,
  // arrow down to "Move up", and activate it. Still 3 slides.
  await page.getByTestId("slide-filmstrip").click();
  await page.getByTestId("slide-filmstrip").focus();
  await page.keyboard.press("ContextMenu");
  await expect(page.getByTestId("slide-context-menu")).toBeVisible();
  await page.keyboard.press("ArrowDown"); // Duplicate
  await page.keyboard.press("ArrowDown"); // Delete
  await page.keyboard.press("ArrowDown"); // Move up
  await page.keyboard.press("Enter");
  await expect(thumbs).toHaveCount(3);

  // Delete one → 2.
  await page.getByTestId("presentation-slide-menu").first().click();
  await page.getByTestId("slide-menu-delete").click();
  await expect(thumbs).toHaveCount(2);

  // Undo the delete → 3 (AC-40/44).
  await page.getByTestId("presentation-undo").click();
  await expect(thumbs).toHaveCount(3);

  // Download the deck and re-open the file — the filmstrip shows 3 slides.
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("presentation-download").click();
  const download = await downloadPromise;
  const filePath = join(tmpdir(), "spec080-deck.pptx");
  await download.saveAs(filePath);

  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(filePath);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("filmstrip-slide")).toHaveCount(3);
});
