import { test, expect, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-108 — the assembly sheet, proven end to end by "Convert to PDF…"
// (AC-50/51): open the library with spec107-shot.png (64 × 40 px) → row
// menu → Convert to PDF… → A4, contain, Narrow margins → the preview reports
// the placed size → Save to library → a new PDF row whose details panel
// reads 1 page at 21 × 29.7 cm; then again with "Fit to content" → the new
// PDF's page is the image's own aspect within 0.01 cm.

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");
const pngFixture = join(fixturesDir, "spec107-shot.png");

async function dropPng(page: Page) {
  const transfer = await page.evaluateHandle(
    (base64) => {
      const dt = new DataTransfer();
      const bytes = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));
      dt.items.add(new File([bytes], "spec107-shot.png", { type: "image/png" }));
      return dt;
    },
    readFileSync(pngFixture).toString("base64"),
  );
  await page.dispatchEvent("body", "drop", { dataTransfer: transfer });
  await expect(page).toHaveURL(/\/media\/.+/);
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();
}

function rowByTitle(page: Page, title: string) {
  return page
    .getByTestId("library-row")
    .filter({ has: page.getByTestId("row-title").getByText(title, { exact: true }) });
}

async function openConvertDialog(page: Page) {
  await rowByTitle(page, "spec107-shot").getByTestId("row-menu").click();
  await page.getByTestId("menu-convert-pdf").click();
  const dialog = page.getByTestId("convert-pdf-dialog");
  await expect(dialog).toBeVisible();
  return dialog;
}

/** Opens a library row in the media viewer and reads its details panel's
 * page size in centimetres ("595 × 842 pt · 21 × 29.7 cm" → [21, 29.7]). */
async function pdfPageSizeCm(page: Page, title: string): Promise<[number, number]> {
  const row = rowByTitle(page, title);
  await expect(row.getByTestId("row-format")).toHaveText("PDF");
  await row.dblclick();
  await expect(page).toHaveURL(/\/media\/.+/);
  await expect(page.getByTestId("detail-pages")).toHaveText("1 page");
  const label = (await page.getByTestId("detail-page-size").textContent()) ?? "";
  const match = /([\d.]+) × ([\d.]+) cm/.exec(label);
  expect(match, label).not.toBeNull();
  return [Number(match![1]), Number(match![2])];
}

test("Convert to PDF: A4 contain with Narrow margins, then Fit to content (AC-50, AC-51)", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("app-bar")).toBeVisible();
  await dropPng(page);
  await expect(rowByTitle(page, "spec107-shot").getByTestId("row-format")).toHaveText("PNG");

  // --- A4, contain, Narrow ---------------------------------------------------
  const dialog = await openConvertDialog(page);
  await expect(dialog.getByTestId("convert-pdf-title")).toHaveValue("spec107-shot");
  await dialog.getByTestId("sheet-paper").selectOption("A4");
  await dialog.getByTestId("sheet-fit").selectOption("contain");
  await dialog.getByTestId("sheet-margin-preset").selectOption("Narrow");
  await expect(dialog.getByTestId("sheet-margin-left")).toHaveValue("1.27");
  // 64 × 40 contained in A4's 18.46 × 27.16 cm content box.
  await expect(dialog.getByTestId("sheet-preview-page-size")).toHaveText("Page 21 × 29.7 cm");
  await expect(dialog.getByTestId("sheet-preview-placed-size")).toHaveText("Placed 18.46 × 11.54 cm");
  await expect(dialog.getByTestId("sheet-estimate")).toContainText("1240 × 1754 px at 150 dpi");
  await dialog.getByTestId("convert-pdf-save").click();
  await expect(dialog).toBeHidden();

  // A new PDF row; the source is untouched.
  await expect(rowByTitle(page, "spec107-shot.pdf")).toHaveCount(1);
  await expect(rowByTitle(page, "spec107-shot").getByTestId("row-format")).toHaveText("PNG");
  const [a4Width, a4Height] = await pdfPageSizeCm(page, "spec107-shot.pdf");
  expect([a4Width, a4Height]).toEqual([21, 29.7]);
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();

  // --- Fit to content ---------------------------------------------------------
  // At 72 dpi one image pixel is one point; no margins, so the page IS the
  // image: 64 × 40 px → 2.258 × 1.411 cm.
  const again = await openConvertDialog(page);
  // The session kept the last sheet.
  await expect(again.getByTestId("sheet-margin-preset")).toHaveValue("Narrow");
  await again.getByTestId("sheet-paper").selectOption("Fit to content");
  await again.getByTestId("sheet-dpi").selectOption("72");
  await expect(again.getByTestId("sheet-margin-link")).toBeChecked();
  await again.getByTestId("sheet-margin-top").fill("0");
  await again.getByTestId("sheet-margin-top").press("Enter");
  await expect(again.getByTestId("sheet-margin-bottom")).toHaveValue("0");
  await expect(again.getByTestId("sheet-landscape")).toBeDisabled();
  await expect(again.getByTestId("sheet-preview-page-size")).toHaveText("Page 2.26 × 1.41 cm");
  await again.getByTestId("convert-pdf-save").click();
  await expect(again).toBeHidden();

  const [width, height] = await pdfPageSizeCm(page, "spec107-shot.pdf (2)");
  expect(Math.abs(width - (64 / 72) * 2.54)).toBeLessThanOrEqual(0.01);
  expect(Math.abs(height - (40 / 72) * 2.54)).toBeLessThanOrEqual(0.01);
  // The image's own aspect: 64 : 40.
  expect(Math.abs(height - (width * 40) / 64)).toBeLessThanOrEqual(0.01);
});

test("the selection bar converts exactly one selected image; a refused margin says why", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("app-bar")).toBeVisible();
  await dropPng(page);

  await rowByTitle(page, "spec107-shot").getByTestId("row-title").click({ modifiers: ["ControlOrMeta"] });
  const bulk = page.getByTestId("bulk-convert-pdf");
  await expect(bulk).toBeEnabled();
  await bulk.click();
  const dialog = page.getByTestId("convert-pdf-dialog");
  await expect(dialog).toBeVisible();

  // Margins that leave less than 1 cm are rejected at input, with the reason.
  await dialog.getByTestId("sheet-paper").selectOption("A4");
  await dialog.getByTestId("sheet-margin-preset").selectOption("Normal");
  await dialog.getByTestId("sheet-margin-top").fill("10.01");
  await dialog.getByTestId("sheet-margin-top").press("Enter");
  await expect(dialog.getByTestId("sheet-error")).toContainText("These margins leave 0.98 × 9.68 cm");
  await expect(dialog.getByTestId("sheet-preview-placed-size")).toHaveText("Placed 15.92 × 9.95 cm");

  // A transparent background is refused for PDF before committing.
  await dialog.getByTestId("sheet-background").selectOption("transparent");
  await expect(dialog.getByTestId("sheet-background-notice")).toHaveText(
    "PDF pages have no transparency, so the background will be white.",
  );

  await dialog.getByTestId("convert-pdf-cancel").click();
  await expect(dialog).toBeHidden();
  await expect(page.getByTestId("library-row")).toHaveCount(1);
});
