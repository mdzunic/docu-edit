import { test, expect, type Download, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-110 — assemble images into a PDF, end to end: open the library with
// three image fixtures (spec107-shot.png, spec107-photo.jpg,
// spec107-anim.gif) → select all → "Combine…" → format PDF, one per page,
// A4 portrait, page numbers on → the strip shows 3 pages → Save to library →
// the new PDF row's details panel reads 3 pages at 21 × 29.7 cm → download
// it, and the JPEG fixture's bytes appear verbatim inside the downloaded
// file (passthrough, proven end to end) → back in the workspace, n-up 4 →
// the strip shows 1 page.

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

async function dropFile(page: Page, name: string, type: string) {
  const transfer = await page.evaluateHandle(
    ({ base64, fileName, mime }) => {
      const dt = new DataTransfer();
      const bytes = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));
      dt.items.add(new File([bytes], fileName, { type: mime }));
      return dt;
    },
    { base64: readFileSync(join(fixturesDir, name)).toString("base64"), fileName: name, mime: type },
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

test("three images → one PDF page each, page numbers on → saved, downloaded with the JPEG verbatim → n-up 4 is one page", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("app-bar")).toBeVisible();
  await dropFile(page, "spec107-shot.png", "image/png");
  await dropFile(page, "spec107-photo.jpg", "image/jpeg");
  await dropFile(page, "spec107-anim.gif", "image/gif");

  // --- select all → Combine… --------------------------------------------------
  await rowByTitle(page, "spec107-shot").getByTestId("row-title").click({ modifiers: ["ControlOrMeta"] });
  await page.keyboard.press("ControlOrMeta+a");
  await expect(page.getByTestId("selection-count")).toHaveText("3 selected");
  await page.getByTestId("bulk-combine").click();
  await expect(page).toHaveURL(/\/assemble$/);
  const items = page.getByTestId("assembly-tray").getByTestId("assembly-item");
  await expect(items).toHaveCount(3);

  // --- PDF, one per page, A4 portrait, page numbers on --------------------------
  await page.getByTestId("assembly-format").selectOption("pdf");
  await expect(page.getByTestId("assembly-page-mode")).toHaveValue("one-per-page");
  await page.getByTestId("assembly-page-mode").selectOption("one-per-page");
  await page.getByTestId("sheet-paper").selectOption("A4");
  await page.getByTestId("sheet-portrait").click();
  await expect(page.getByTestId("sheet-portrait")).toHaveAttribute("aria-pressed", "true");
  const numbers = page.getByTestId("assembly-page-numbers");
  await expect(numbers).not.toBeChecked();
  await numbers.check();

  // The strip shows 3 pages, the first current; the count says so too.
  const strip = page.getByTestId("assembly-pages").getByTestId("assembly-page");
  await expect(strip).toHaveCount(3);
  await expect(page.getByTestId("assembly-page-count")).toHaveText("3 pages");
  await expect(strip.nth(0)).toHaveAttribute("aria-current", "page");
  await strip.nth(1).click();
  await expect(strip.nth(1)).toHaveAttribute("aria-current", "page");
  await page.keyboard.press("PageDown");
  await expect(strip.nth(2)).toHaveAttribute("aria-current", "page");
  await expect(page.getByTestId("assembly-preview")).toHaveAttribute("aria-label", /^Preview of page 3 of 3: 21 × 29\.7 cm/);

  // The photo is a baseline JPEG, uncropped: it rides in as the original.
  const photoRow = items.filter({ has: page.getByTestId("item-title").getByText("spec107-photo", { exact: true }) });
  await expect(photoRow.getByTestId("item-encoding")).toHaveText("original JPEG");
  const shotRow = items.filter({ has: page.getByTestId("item-title").getByText("spec107-shot", { exact: true }) });
  await expect(shotRow.getByTestId("item-encoding")).toHaveText("re-encoded: PNG");
  await expect(page.getByTestId("assembly-estimate")).toContainText("3 pages · PDF up to about");

  // --- Save to library ------------------------------------------------------
  await page.getByTestId("assembly-save").click();
  await expect(page.getByTestId("assembly-saved")).toHaveText("Saved “Combined image” to the library.", { timeout: 30_000 });
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();

  const pdfRow = rowByTitle(page, "Combined image");
  await expect(pdfRow.getByTestId("row-format")).toHaveText("PDF");
  // The sources are untouched.
  await expect(rowByTitle(page, "spec107-photo").getByTestId("row-format")).toHaveText("JPEG");
  await pdfRow.dblclick();
  await expect(page).toHaveURL(/\/media\/.+/);
  await expect(page.getByTestId("detail-pages")).toHaveText("3 pages");
  const size = (await page.getByTestId("detail-page-size").textContent()) ?? "";
  expect(size, size).toMatch(/\b21 × 29\.7 cm\b/);
  expect(size).not.toContain("mixed");
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();

  // --- download it: the JPEG fixture's own bytes are inside ---------------------
  const downloads: Download[] = [];
  page.on("download", (download) => downloads.push(download));
  await rowByTitle(page, "Combined image").getByTestId("row-select").check();
  await expect(page.getByTestId("selection-count")).toHaveText("1 selected");
  await page.getByTestId("bulk-download").click();
  await expect.poll(() => downloads.length, { timeout: 15_000 }).toBe(1);
  expect(downloads[0].suggestedFilename()).toBe("Combined image.pdf");
  const downloadedPath = await downloads[0].path();
  expect(downloadedPath).toBeTruthy();
  const pdf = readFileSync(downloadedPath!);
  expect(pdf.subarray(0, 5).toString("latin1")).toBe("%PDF-");
  const jpeg = readFileSync(join(fixturesDir, "spec107-photo.jpg"));
  expect(pdf.indexOf(jpeg)).toBeGreaterThan(0);

  // --- back in the workspace: n-up 4 is one page ----------------------------------
  await page.goto("/assemble");
  await expect(items).toHaveCount(3);
  await expect(page.getByTestId("assembly-format")).toHaveValue("pdf");
  await expect(strip).toHaveCount(3);
  await page.getByTestId("assembly-page-mode").selectOption("n-up");
  await page.getByTestId("assembly-n-up").selectOption("4");
  await expect(strip).toHaveCount(1);
  await expect(page.getByTestId("assembly-page-count")).toHaveText("1 page");
  await expect(page.getByTestId("assembly-placed-item")).toHaveCount(3);
});
