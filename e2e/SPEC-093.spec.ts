import { test, expect, type Page } from "@playwright/test";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { renderDateTime } from "../src/lib/fields";
import { DEFAULT_LANGUAGE } from "../src/lib/core/language";

// SPEC-093 — slide footers, slide numbers and dates. Open a deck whose second
// slide shows a footer, a live date and the number "2"; configure the deck
// through the Header & Footer dialog and Apply to all; confirm the title slide
// stays blank (Don't show on title slide); duplicate slide 2 and watch the
// numbers renumber live; download → reopen and confirm the three values render,
// the date being today's (no clock is pinned — SPEC-053's e2e precedent).
const here = dirname(fileURLToPath(import.meta.url));
const PPTX = join(here, "fixtures", "spec093-headers-footers.pptx");

async function openFixture(page: Page, file: string): Promise<void> {
  await page.goto("/");
  await page.getByTestId("open-document-input").setInputFiles(file);
  await expect(page).toHaveURL(/\/slides\/.+/);
  await expect(page.getByTestId("presentation-view")).toBeVisible();
}

async function selectSlide(page: Page, index: number): Promise<void> {
  await page.getByTestId("filmstrip-slide").nth(index).click();
}

function stageField(page: Page, kind: "footer" | "slide-number" | "date") {
  return page.getByTestId("stage-workspace").locator(`[data-testid="slide-field"][data-field-kind="${kind}"]`);
}

async function downloadTo(page: Page, name: string): Promise<string> {
  const downloadPromise = page.waitForEvent("download");
  await page.getByTestId("presentation-download").click();
  const download = await downloadPromise;
  const filePath = join(tmpdir(), name);
  await download.saveAs(filePath);
  return filePath;
}

test("pptx: footers, live numbers and the date round-trip through the dialog", async ({ page }) => {
  await openFixture(page, PPTX);

  // AC-107: slide 2 (index 1) shows a footer, the date and the number "2".
  await selectSlide(page, 1);
  await expect(stageField(page, "footer")).toBeVisible();
  await expect(stageField(page, "date")).toBeVisible();
  await expect(stageField(page, "slide-number")).toHaveText("2");

  // AC-108: Header & footer → set footer text, tick slide number + auto date,
  // Apply to all.
  await page.getByTestId("slides-header-footer").click();
  const dialog = page.getByTestId("header-footer-dialog");
  await expect(dialog).toBeVisible();
  await expect(dialog.getByRole("heading", { name: "Header and footer" })).toBeVisible();
  await page.getByTestId("footer-show").check();
  await page.getByTestId("footer-text").fill("Confidential");
  await page.getByTestId("footer-slide-number").check();
  await page.getByTestId("footer-date-show").check();
  await page.getByTestId("footer-date-auto").check();
  await page.getByTestId("footer-apply-all").click();
  await expect(dialog).toBeHidden();

  // AC-109: slide 1 (the title slide, index 0) still shows nothing — the deck
  // has "Don't show on title slide" set.
  await selectSlide(page, 0);
  await expect(page.getByTestId("stage-workspace").locator('[data-testid="slide-field"]')).toHaveCount(0);

  // AC-110: duplicate slide 2 → the copy reads "3" and the following reads "4",
  // with no further input (the numbers are live from index).
  await selectSlide(page, 1);
  await page.getByTestId("presentation-slide-menu").nth(1).click();
  await page.getByTestId("slide-menu-duplicate").click();
  await selectSlide(page, 2); // the duplicate
  await expect(stageField(page, "slide-number")).toHaveText("3");
  await selectSlide(page, 3); // the slide that followed
  await expect(stageField(page, "slide-number")).toHaveText("4");

  // AC-111: download → reopen → the same three values render, the date being
  // today's under the default pattern.
  const file = await downloadTo(page, "spec093-roundtrip.pptx");
  await openFixture(page, file);
  await selectSlide(page, 1);
  await expect(stageField(page, "footer")).toHaveText("Confidential");
  await expect(stageField(page, "slide-number")).toHaveText("2");
  const expectedDate = renderDateTime("d MMMM yyyy", new Date(), DEFAULT_LANGUAGE);
  await expect(stageField(page, "date")).toHaveText(expectedDate ?? "");
});
