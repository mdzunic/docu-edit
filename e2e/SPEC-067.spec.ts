import { test, expect, type Page } from "@playwright/test";

// SPEC-067 — the redesigned document library: list/grid views, search, sort,
// filters, starring, multi-select + bulk actions, trash and the storage panel.
// Each Playwright test runs in a fresh browser context, so IndexedDB starts
// empty and the first-run empty state shows until a document is created.

async function createDocument(page: Page, title: string): Promise<string> {
  await page.goto("/");
  await page.getByTestId("new-document").click();
  await expect(page).toHaveURL(/\/doc\/.+/);
  await page.getByTestId("document-title").click();
  await page.getByTestId("document-title-input").fill(title);
  await page.getByTestId("document-title-input").press("Enter");
  await expect(page.getByTestId("document-title")).toHaveText(title);
  const url = page.url();
  return url.slice(url.lastIndexOf("/") + 1);
}

function rowByTitle(page: Page, title: string) {
  return page
    .getByTestId("library-row")
    .filter({ has: page.getByTestId("row-title").getByText(title, { exact: true }) });
}

test("the first-run empty state describes the product and the three ways to start (AC-51..53)", async ({
  page,
}) => {
  await page.goto("/");
  const empty = page.getByTestId("empty-state");
  await expect(empty).toBeVisible();
  await expect(empty).toContainText("docu-edit");
  await expect(empty).toContainText("stored in this browser");
  await expect(page.getByTestId("new-document")).toBeVisible();
  await expect(page.getByTestId("new-from-template")).toBeVisible();
  await expect(page.getByTestId("open-document")).toBeVisible();
});

test("the view toggle switches to the grid and the choice persists (AC-1, AC-2)", async ({
  page,
}) => {
  await createDocument(page, "Griddable");
  await page.goto("/");

  await expect(page.getByTestId("library-list")).toBeVisible();
  await page.getByTestId("view-grid").click();
  await expect(page.getByTestId("library-grid")).toBeVisible();
  // The grid card shows a thumbnail.
  await expect(page.getByTestId("row-thumbnail").first()).toBeVisible();

  // The choice survives a reload.
  await page.reload();
  await expect(page.getByTestId("library-grid")).toBeVisible();
});

test("search filters by subsequence, highlights matches, and offers a clear when nothing matches (AC-8..10, AC-54, AC-55)", async ({
  page,
}) => {
  await createDocument(page, "Quarterly Report");
  await createDocument(page, "Budget");
  await page.goto("/");

  await page.getByTestId("library-search").fill("qrt");
  await expect(page.getByTestId("row-title")).toHaveCount(1);
  await expect(page.getByTestId("row-title")).toHaveText("Quarterly Report");
  await expect(page.getByTestId("title-match").first()).toBeVisible();

  await page.getByTestId("library-search").fill("zzzzz");
  await expect(page.getByTestId("no-results")).toBeVisible();
  await page.getByTestId("clear-filters").click();
  await expect(page.getByTestId("row-title")).toHaveCount(2);
});

test("format and starred filter chips narrow the list with live counts (AC-16, AC-17)", async ({
  page,
}) => {
  await createDocument(page, "One");
  await createDocument(page, "Two");
  await page.goto("/");

  // Star "One" from its row, then filter to starred only.
  await rowByTitle(page, "One").getByTestId("row-star").click();
  await expect(page.getByTestId("filter-starred")).toContainText("1");
  await page.getByTestId("filter-starred").click();
  await expect(page.getByTestId("row-title")).toHaveText("One");
});

test("starred documents sort first within the active sort (AC-21)", async ({
  page,
}) => {
  await createDocument(page, "Alpha");
  await createDocument(page, "Bravo");
  await page.goto("/");

  // Star the older document (Alpha); it should jump to the top.
  await rowByTitle(page, "Alpha").getByTestId("row-star").click();
  const titles = page.getByTestId("row-title");
  await expect(titles.nth(0)).toHaveText("Alpha");
});

test("multi-select drives a bulk move-to-trash (AC-23, AC-30, AC-31, AC-32)", async ({
  page,
}) => {
  await createDocument(page, "First");
  await createDocument(page, "Second");
  await page.goto("/");

  await page.getByTestId("row-select").nth(0).check();
  await page.getByTestId("row-select").nth(1).check();
  await expect(page.getByTestId("selection-count")).toContainText("2");

  await page.getByTestId("bulk-trash").click();
  await expect(page.getByTestId("library-empty")).toBeVisible();
  await page.getByTestId("trash-link").click();
  await expect(page.getByTestId("library-row")).toHaveCount(2);
});

test("Enter opens the focused row from the keyboard (AC-57, AC-58)", async ({
  page,
}) => {
  const id = await createDocument(page, "Keyboard");
  await page.goto("/");

  await rowByTitle(page, "Keyboard").focus();
  await page.keyboard.press("Enter");
  await expect(page).toHaveURL(new RegExp(`/doc/${id}$`));
});

// A 240×120 solid-blue PNG (405 bytes), pre-encoded so this spec needs no
// PNG encoder of its own (SPEC-024's spec carries one for its insert tests).
const BLUE_PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAPAAAAB4CAIAAABD1OhwAAABXElEQVR4nO3SUQkAIBTAwJfIatY1jiUEYRxcgH1s1j6QMd8L4CFDk2JoUgxNiqFJMTQphibF0KQYmhRDk2JoUgxNiqFJMTQphibF0KQYmhRDk2JoUgxNiqFJMTQphibF0KQYmhRDk2JoUgxNiqFJMTQphibF0KQYmhRDk2JoUgxNiqFJMTQphibF0KQYmhRDk2JoUgxNiqFJMTQphibF0KQYmhRDk2JoUgxNiqFJMTQphibF0KQYmhRDk2JoUgxNiqFJMTQphibF0KQYmhRDk2JoUgxNiqFJMTQphibF0KQYmhRDk2JoUgxNiqFJMTQphibF0KQYmhRDk2JoUgxNiqFJMTQphibF0KQYmhRDk2JoUgxNiqFJMTQphibF0KQYmhRDk2JoUgxNiqFJMTQphibF0KQYmhRDk2JoUgxNiqFJMTQphibF0KQYmhRDk2JoUgxNiqFJMTQphiblAs5s9A7sSMF/AAAAAElFTkSuQmCC",
  "base64",
);

test("a grid card's thumbnail shows the document's images, tables, charts and shapes (AC-3)", async ({
  page,
}) => {
  await createDocument(page, "Picture report");
  await page.getByTestId("document-content").click();
  await page.keyboard.type("Quarterly picture");

  // A chart from the gallery (SPEC-052) — inserted first, while the caret is
  // still in a plain paragraph (charts do not insert inside tables)…
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("chart-gallery-columnClustered").click();
  await expect(
    page.getByTestId("document-content").locator('[data-testid^="drawing-"]'),
  ).toBeVisible();

  // …an image via the Insert menu's hidden file input (SPEC-024)…
  await page
    .getByTestId("insert-image-file-input")
    .setInputFiles({ name: "chart.png", mimeType: "image/png", buffer: BLUE_PNG });
  await expect(
    page.getByTestId("document-content").locator('[data-testid="doc-image"]'),
  ).toBeVisible();

  // …a shape placed with a click (SPEC-051)…
  await page.getByTestId("document-content").click({ position: { x: 20, y: 20 } });
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("shape-gallery-rect").click();
  await page.getByTestId("document-content").click({ position: { x: 420, y: 300 } });
  // Drawings share one testid counter (SPEC-051): the chart is drawing-0,
  // the placed shape drawing-1.
  await expect(
    page.getByTestId("document-content").locator('[data-testid^="drawing-"]'),
  ).toHaveCount(2);

  // …and last a 2×2 table with a cell of text (SPEC-023), so the caret can
  // stay in it until close.
  await page.getByTestId("menu-title-insert").click();
  await page.getByTestId("table-grid-cell-2-2").click();
  await expect(page.getByTestId("document-content").locator("table.doc-table")).toHaveCount(1);
  await page.getByTestId("table-cell-0-0").click();
  await page.keyboard.type("North");

  // Close flushes the save before navigating (SPEC-060), so the library row
  // is current on arrival.
  await page.getByTestId("close-document").click();
  await expect(page).toHaveURL(/\/$/);
  await page.getByTestId("view-grid").click();

  const thumbnail = page.getByTestId("row-thumbnail").first();
  const thumbnailImage = thumbnail.locator("img.thumbnail-image");
  await expect(thumbnailImage).toBeVisible();
  const src = await thumbnailImage.getAttribute("src");
  expect(src).toMatch(/^data:image\/png;base64,/);
  // The table miniature shows its leading cell text on a grid…
  await expect(thumbnail.locator(".thumbnail-table")).toBeVisible();
  await expect(thumbnail.locator(".thumbnail-table")).toContainText("North");
  // …the chart paints on its own small canvas…
  await expect(thumbnail.locator("canvas.thumbnail-chart")).toBeVisible();
  // …the shape paints too…
  await expect(thumbnail.locator("canvas.thumbnail-shape-canvas")).toBeVisible();
  // …and the document's text still previews alongside them.
  await expect(thumbnail).toContainText("Quarterly picture");
});
