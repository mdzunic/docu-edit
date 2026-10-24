import { test, expect, type Download, type Page } from "@playwright/test";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

// SPEC-111 — merge PDFs and organise their pages, end to end: open the
// library with spec107-pages.pdf (3 pages: Letter, A4 at /Rotate 90, and
// 720 × 540 pt) and the 2-page spec111-outlines.pdf (A5 portrait, then
// 595 × 420 pt, neither rotated; a tagged PDF) → select both, the outlines
// first → "Combine…" → format PDF → the tray shows 5 page cards → rotate
// page 2 right, delete page 4, drag page 5 to the front → the result preview
// is the browser's own viewer → Save to library → "What changed" lists the
// dropped structure tree → the new row's details panel reads 4 pages, page 1
// the former page 5's size and page 3 (the former page 2) "rotate 90" →
// download it and re-open it → the same 4 pages.

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "fixtures");

// Tall enough that the tray shows every card at once: a drag between two
// cards needs both on screen, or Playwright's scroll to the target between
// mousedown and the first move starts the drag on whichever card has
// scrolled under the pointer.
test.use({ viewport: { width: 1280, height: 1600 } });

async function dropBytes(page: Page, bytes: Buffer, fileName: string) {
  const transfer = await page.evaluateHandle(
    ({ base64, name }) => {
      const dt = new DataTransfer();
      const data = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));
      dt.items.add(new File([data], name, { type: "application/pdf" }));
      return dt;
    },
    { base64: bytes.toString("base64"), name: fileName },
  );
  await page.dispatchEvent("body", "drop", { dataTransfer: transfer });
  await expect(page).toHaveURL(/\/media\/.+/);
}

async function dropFixture(page: Page, name: string) {
  await dropBytes(page, readFileSync(join(fixturesDir, name)), name);
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();
}

function rowByTitle(page: Page, title: string) {
  return page
    .getByTestId("library-row")
    .filter({ has: page.getByTestId("row-title").getByText(title, { exact: true }) });
}

/** The details panel's line for each page. */
async function expectPages(page: Page) {
  await expect(page.getByTestId("detail-pages")).toHaveText("4 pages");
  const lines = page.getByTestId("detail-page-list").getByTestId("detail-page");
  await expect(lines).toHaveCount(4);
  // Page 1 is the former page 5: spec107-pages' 720 × 540 pt page.
  await expect(lines.nth(0)).toContainText("Page 1 · 720 × 540 pt");
  await expect(lines.nth(0)).not.toContainText("rotate");
  // Page 2 is the outlines' first page; page 3 the former page 2, turned.
  await expect(lines.nth(1)).toContainText("Page 2 · 420 × 595 pt");
  await expect(lines.nth(2)).toContainText("Page 3 · 595 × 420 pt");
  await expect(lines.nth(2)).toContainText("rotate 90");
  // Page 4 is spec107-pages' first, US Letter; its A4 page was deleted.
  await expect(lines.nth(3)).toContainText("Page 4 · 612 × 792 pt");
  await expect(page.getByTestId("detail-page-list")).not.toContainText("595 × 842");
}

test("two PDFs → 5 page cards → rotate, delete, drag → saved as 4 pages with what changed → downloaded and re-opened", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("app-bar")).toBeVisible();
  await dropFixture(page, "spec107-pages.pdf");
  await dropFixture(page, "spec111-outlines.pdf");

  // --- select both, the outlines first → Combine… --------------------------------
  const combine = page.getByTestId("bulk-combine");
  await rowByTitle(page, "spec111-outlines").getByTestId("row-title").click({ modifiers: ["ControlOrMeta"] });
  await expect(page.getByTestId("selection-count")).toHaveText("1 selected");
  // One PDF is already something to organise.
  await expect(combine).toBeEnabled();
  await rowByTitle(page, "spec107-pages").getByTestId("row-title").click({ modifiers: ["ControlOrMeta"] });
  await expect(page.getByTestId("selection-count")).toHaveText("2 selected");
  await combine.click();
  await expect(page).toHaveURL(/\/assemble$/);

  // --- format PDF: the tray shows 5 page cards ---------------------------------
  await page.getByTestId("assembly-format").selectOption("pdf");
  await expect(page.getByTestId("assembly-format")).toHaveValue("pdf");
  const cards = page.getByTestId("assembly-tray").getByTestId("assembly-page-card");
  await expect(cards).toHaveCount(5);
  await expect(cards.getByTestId("page-source")).toHaveText([
    "spec111-outlines · p. 1 of 2",
    "spec111-outlines · p. 2 of 2",
    "spec107-pages · p. 1 of 3",
    "spec107-pages · p. 2 of 3",
    "spec107-pages · p. 3 of 3",
  ]);
  await expect(cards.getByTestId("page-size")).toHaveText(["420 × 595 pt", "595 × 420 pt", "612 × 792 pt", "595 × 842 pt", "720 × 540 pt"]);
  await expect(cards.getByTestId("page-rotation")).toHaveText(["rotate 0", "rotate 0", "rotate 0", "rotate 90", "rotate 0"]);
  // No faked preview: each card is paper at its page's aspect, and says so.
  await expect(cards.first()).toContainText("No preview");
  await expect(cards.first().locator("img, canvas")).toHaveCount(0);
  const aspect = Number(await cards.nth(4).getByTestId("page-paper").getAttribute("data-aspect"));
  expect(aspect).toBeCloseTo(720 / 540, 3);

  // --- rotate page 2 right ------------------------------------------------------------
  await cards.nth(1).getByTestId("page-rotate-right").click();
  await expect(cards.nth(1).getByTestId("page-rotation")).toHaveText("rotate 90");

  // --- delete page 4 (from the plan) ------------------------------------------------
  await cards.nth(3).getByTestId("page-delete").click();
  await expect(cards).toHaveCount(4);
  await expect(page.getByTestId("assembly-tray-status")).toHaveText("Removed page 2 of “spec107-pages” from the plan.");

  // --- drag page 5 (now the fourth card) to the front ----------------------------
  await cards.nth(3).dragTo(cards.nth(0), { sourcePosition: { x: 20, y: 20 }, targetPosition: { x: 20, y: 12 } });
  await expect(cards.getByTestId("page-source")).toHaveText([
    "spec107-pages · p. 3 of 3",
    "spec111-outlines · p. 1 of 2",
    "spec111-outlines · p. 2 of 2",
    "spec107-pages · p. 1 of 3",
  ]);
  await expect(cards.getByTestId("page-rotation")).toHaveText(["rotate 0", "rotate 0", "rotate 90", "rotate 0"]);

  // --- the result preview is the browser's own viewer over the planned PDF ---------
  await expect(page.getByTestId("merge-preview").locator("object")).toHaveAttribute("data", /^blob:/, { timeout: 15_000 });
  await expect(page.getByTestId("merge-refresh")).toBeVisible();

  // --- Save to library ----------------------------------------------------------------
  await expect(page.getByTestId("merge-destination")).toHaveValue("merge");
  await page.getByTestId("assembly-save").click();
  await expect(page.getByTestId("assembly-saved")).toHaveText("Saved “Combined PDF” to the library.", { timeout: 30_000 });

  // --- What changed: the tagged source's structure tree was not carried ------------
  const drops = page.getByTestId("merge-drops");
  await expect(drops).toBeVisible();
  await drops.locator("summary").click();
  const structure = drops.locator('[data-testid="merge-drop-group"][data-kind="structureTree"]');
  await expect(structure).toBeVisible();
  await expect(structure.getByTestId("merge-drop-count")).toHaveText("(1)");
  await expect(structure.getByTestId("merge-drop")).toContainText("“spec111-outlines”: Its structure tree");

  // --- the new row's details panel ------------------------------------------------------
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();
  const merged = rowByTitle(page, "Combined PDF");
  await expect(merged.getByTestId("row-format")).toHaveText("PDF");
  // The sources are untouched.
  await rowByTitle(page, "spec107-pages").dblclick();
  await expect(page.getByTestId("detail-pages")).toHaveText("3 pages");
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();
  await merged.dblclick();
  await expect(page).toHaveURL(/\/media\/.+/);
  await expectPages(page);
  await page.getByTestId("close-document").click();
  await expect(page.getByTestId("library")).toBeVisible();

  // --- download it, and re-open it: the same 4 pages ---------------------------------
  const downloads: Download[] = [];
  page.on("download", (download) => downloads.push(download));
  await rowByTitle(page, "Combined PDF").getByTestId("row-select").check();
  await expect(page.getByTestId("selection-count")).toHaveText("1 selected");
  await page.getByTestId("bulk-download").click();
  await expect.poll(() => downloads.length, { timeout: 15_000 }).toBe(1);
  expect(downloads[0].suggestedFilename()).toBe("Combined PDF.pdf");
  const downloadedPath = await downloads[0].path();
  expect(downloadedPath).toBeTruthy();
  const bytes = readFileSync(downloadedPath!);
  expect(bytes.subarray(0, 8).toString("latin1")).toBe("%PDF-1.7");
  await page.getByTestId("clear-selection").click();
  await dropBytes(page, bytes, "Reopened merge.pdf");
  await expectPages(page);
});
